// api/mail-new.js
/**
 * Vercel Serverless 安全版（覆盖此文件即可）
 * - 仅用 IMAP OAuth2（你的 refresh_token 属于 IMAP/SMTP，不走 Graph）
 * - 按 UID 抓最新一封；超过阈值（默认 1.8MB）不解析，返回提示 + 下载/源码按钮
 * - ?response_type=html|json ；?raw=1（前60KB），?download=1（整封 eml）
 */

const Imap = require('node-imap');
const { simpleParser } = require('mailparser');
const fetch = require('node-fetch');

const MAX_HTML_SIZE = parseInt(process.env.MAX_HTML_SIZE || '1800000', 10); // ~1.8MB
const MAX_RAW_BUFFER = parseInt(process.env.MAX_RAW_BUFFER || '5242880', 10); // 5MB，超过就直接提示（避免超时/内存）

function escapeHtml(s=''){return String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function renderHtmlPage({from,subject,date,htmlBody,textBody,hint}){
  const htmlOrText = htmlBody ? htmlBody : `<pre style="white-space:pre-wrap;">${escapeHtml(textBody||'')}</pre>`;
  const notice = hint ? `<div class="warn">${hint}</div>` : '';
  return `<!doctype html><html><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>邮件信息</title>
<style>
body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Noto Sans,Arial;background:#f7f7f7;margin:0}
.wrap{max-width:960px;margin:24px auto;background:#fff;border-radius:12px;box-shadow:0 10px 24px rgba(0,0,0,.06);overflow:hidden}
.hdr{padding:18px 20px;border-bottom:1px solid #eee}.hdr h1{margin:0;font-size:20px}
.meta{padding:10px 20px;border-bottom:1px solid #f5f5f5;color:#555;font-size:14px}
.actions{padding:10px 20px;border-bottom:1px solid #f5f5f5}
.btn{display:inline-block;margin-right:8px;padding:6px 10px;border:1px solid #ddd;border-radius:6px;text-decoration:none;color:#222;background:#fafafa}
.content{padding:16px 20px;min-height:300px}.label{font-weight:600;color:#333}
.warn{padding:12px;background:#fff3cd;border:1px solid #ffeeba;border-radius:8px;color:#856404;margin-bottom:12px}
</style></head><body>
<div class="wrap">
  <div class="hdr"><h1>邮件信息</h1></div>
  <div class="meta">
    <div><span class="label">发件人：</span>${escapeHtml(from||'')}</div>
    <div><span class="label">主题：</span>${escapeHtml(subject||'')}</div>
    <div><span class="label">日期：</span>${escapeHtml(String(date||''))}</div>
  </div>
  <div class="actions">
    <a class="btn" href="?download=1">下载原始邮件(.eml)</a>
    <a class="btn" href="?raw=1" target="_blank">查看源码片段</a>
  </div>
  <div class="content">
    ${notice}
    ${htmlOrText || '<div class="warn">⚠️ 此邮件没有可显示的正文或仅包含附件/图片。</div>'}
  </div>
</div>
</body></html>`;
}

async function oauthAccessToken(refresh_token, client_id){
  const r = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token',{
    method:'POST',
    headers:{'Content-Type':'application/x-www-form-urlencoded'},
    body:new URLSearchParams({ client_id, grant_type:'refresh_token', refresh_token }).toString()
  });
  if(!r.ok) throw new Error(`token error ${r.status}: ${await r.text()}`);
  const j = await r.json();
  if(!j.access_token) throw new Error('no access_token');
  return j.access_token;
}
const genXOAUTH2=(user,token)=>Buffer.from(`user=${user}\x01auth=Bearer ${token}\x01\x01`).toString('base64');

function findBoxName(boxes,want){
  const flat=[]; (function walk(obj,p=''){Object.keys(obj||{}).forEach(n=>{const b=obj[n];const path=p?`${p}${b.delimiter}${n}`:n;flat.push(path);if(b.children)walk(b.children,path);});})(boxes);
  const w=String(want||'INBOX').toLowerCase();
  let hit=flat.find(n=>n.toLowerCase()===w); if(hit) return hit;
  if(w==='junk'){ hit=flat.find(n=>/^(junk|junk[-\s]?email|垃圾邮件|垃圾)$/i.test(n)); if(hit) return hit; }
  if(w==='inbox'){ hit=flat.find(n=>/^inbox$/i.test(n))||'INBOX'; return hit; }
  return flat.find(n=>/^inbox$/i.test(n))||'INBOX';
}

module.exports = async (req,res)=>{
  try{
    const q = req.method==='GET'?req.query:req.body;
    const { refresh_token, client_id, email, mailbox='INBOX', response_type='html', raw='0', download='0' } = q;

    if(!refresh_token || !client_id || !email){
      return res.status(400).json({error:'Missing required parameters: refresh_token, client_id, email'});
    }

    // 1) 换 IMAP 的 access_token
    const token = await oauthAccessToken(refresh_token, client_id);

    const imap = new Imap({
      user: email,
      xoauth2: genXOAUTH2(email, token),
      host: 'outlook.office365.com',
      port: 993,
      tls: true,
      tlsOptions: { rejectUnauthorized: false }
    });

    imap.once('ready', async ()=>{
      try{
        // 2) 定位文件夹并打开（只读）
        const boxes = await new Promise((resolve,reject)=>imap.getBoxes((e,b)=>e?reject(e):resolve(b)));
        const boxName = findBoxName(boxes, mailbox);
        await new Promise((resolve,reject)=>imap.openBox(boxName, true, (e)=>e?reject(e):resolve()));

        // 3) 最新一封（UID）
        const uids = await new Promise((resolve,reject)=>imap.search(['ALL'], (e, ids)=> e?reject(e):resolve(ids.slice(-1))));
        if(!uids || !uids.length){
          const page = renderHtmlPage({from:'',subject:'',date:'',htmlBody:'',textBody:'（此目录暂无邮件）'});
          return res.status(200).type('text/html').send(page);
        }

        // 4) 抓整封（按 UID）；**Vercel 安全：限制最大缓冲**
        const f = imap.fetch(uids, { bodies: '', struct: true, uid: true });

        f.on('message',(msg)=>{
          let rawBuf = '';
          let truncated = false;

          msg.on('body',(stream)=>{
            stream.on('data', chunk=>{
              if (truncated) return;
              rawBuf += chunk.toString('utf8');
              if (rawBuf.length > MAX_RAW_BUFFER) {
                truncated = true; // 超 5MB 直接截断，避免 serverless 超时/爆内存
              }
            });
            stream.once('end', async ()=>{
              try{
                if (String(download)==='1') {
                  res.setHeader('Content-Type','message/rfc822');
                  res.setHeader('Content-Disposition','attachment; filename="message.eml"');
                  return res.status(200).send(rawBuf);
                }
                if (String(raw)==='1') {
                  return res.status(200).type('text/plain').send(rawBuf.slice(0, 60000));
                }

                if (truncated) {
                  // 大邮件：不做解析，直接返回提示页面
                  const page = renderHtmlPage({
                    from:'', subject:'', date:'',
                    htmlBody:'',
                    textBody:'',
                    hint:`⚠️ 邮件体积较大（超过 ${(MAX_RAW_BUFFER/1024/1024).toFixed(1)} MB），
                          为避免 Vercel 函数超时/爆内存，已停止在线解析。
                          你可以点击上方“下载原始邮件(.eml)”完整查看，或用 &raw=1 看源码片段。`
                  });
                  return res.status(200).type('text/html').send(page);
                }

                // 普通邮件：完整解析
                const mail = await simpleParser(rawBuf);
                let html = mail.html || '';
                const text = mail.text || '';
                const textAsHtml = mail.textAsHtml || (text ? `<pre style="white-space:pre-wrap;">${escapeHtml(text)}</pre>` : '');
                let hint = '';

                if (html && html.length > MAX_HTML_SIZE) {
                  hint = `⚠️ 正文较大（${(html.length/1024/1024).toFixed(2)} MB），已切换为简化视图，以避免 Vercel 截断。`;
                  html = textAsHtml || '<div class="warn">正文过大且无纯文本可用，建议下载原文查看。</div>';
                } else if (!html) {
                  html = textAsHtml;
                }

                const data = {
                  send: mail?.from?.text || '',
                  subject: mail.subject || '',
                  date: mail.date || '',
                  html: html || '',
                  text: text || ''
                };

                if (String(response_type).toLowerCase()==='json') {
                  return res.status(200).json(data);
                }

                const page = renderHtmlPage({
                  from: data.send, subject: data.subject, date: data.date,
                  htmlBody: data.html, textBody: data.text, hint
                });
                res.status(200).type('text/html').send(page);
              }catch(e){
                res.status(500).type('text/plain').send('parse error: ' + String(e.message||e));
              }
            });
          });
        });

        f.once('end', ()=> imap.end());
      }catch(e){
        imap.end();
        res.status(500).json({error:String(e.message||e)});
      }
    });

    imap.once('error',(err)=> res.status(500).json({error:String(err.message||err)}));
    imap.connect();

  }catch(e){
    res.status(500).json({error:String(e.message||e)});
  }
};
