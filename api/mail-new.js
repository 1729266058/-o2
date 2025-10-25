/**
 * mail-new.js —— IMAP OAuth2 稳定版（覆盖即可）
 * 关键修复：
 *  - search() 返回 UID；fetch() 必须加 { uid: true }，否则会抓错信/抓不到正文 → 页面空白
 *  - 先把整封原文缓冲完，在 'end' 后交给 mailparser 解析（避免流还没读完就渲染）
 *  - 优先 html，退 text；若都没有，给出黄条提示而不是白板
 *  - /api 格式保持不变：response_type=html|json；支持 &debug=1、&raw=1
 */

const Imap = require('node-imap');
const { simpleParser } = require('mailparser');
const fetch = require('node-fetch'); // Node18 可省，但保留兼容更稳

// ---------- 小工具 ----------
function escapeHtml(s=''){return String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function renderHtmlPage({from,subject,date,htmlBody,textBody}){
  const htmlOrText = htmlBody
    ? htmlBody
    : `<pre style="white-space:pre-wrap;">${escapeHtml(textBody||'')}</pre>`;
  return `<!doctype html><html><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>邮件信息</title>
<style>
body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Noto Sans,Arial;background:#f7f7f7;margin:0}
.wrap{max-width:960px;margin:24px auto;background:#fff;border-radius:12px;box-shadow:0 10px 24px rgba(0,0,0,.06);overflow:hidden}
.hdr{padding:18px 20px;border-bottom:1px solid #eee}.hdr h1{margin:0;font-size:20px}
.meta{padding:10px 20px;border-bottom:1px solid #f5f5f5;color:#555;font-size:14px}
.content{padding:16px 20px;min-height:300px}.label{font-weight:600;color:#333}
.warn{padding:12px;background:#fff3cd;border:1px solid #ffeeba;border-radius:8px;color:#856404}
</style></head><body>
<div class="wrap">
  <div class="hdr"><h1>邮件信息</h1></div>
  <div class="meta">
    <div><span class="label">发件人：</span>${escapeHtml(from||'')}</div>
    <div><span class="label">主题：</span>${escapeHtml(subject||'')}</div>
    <div><span class="label">日期：</span>${escapeHtml(String(date||''))}</div>
    <div style="margin-top:8px" class="label">内容：</div>
  </div>
  <div class="content">${htmlOrText || '<div class="warn">⚠️ 此邮件没有可显示的正文或仅包含附件/图片。</div>'}</div>
</div>
</body></html>`;
}

async function get_access_token(refresh_token, client_id){
  const r = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token',{
    method:'POST',
    headers:{'Content-Type':'application/x-www-form-urlencoded'},
    body:new URLSearchParams({
      client_id, grant_type:'refresh_token', refresh_token
      // 不加 scope：沿用该 refresh_token 绑定的 IMAP/SMTP 资源
    }).toString()
  });
  if(!r.ok){throw new Error(`token error ${r.status}: ${await r.text()}`);}
  const j = await r.json();
  if(!j.access_token) throw new Error('no access_token');
  return j.access_token;
}
const genXOAUTH2=(user,token)=>Buffer.from(`user=${user}\x01auth=Bearer ${token}\x01\x01`).toString('base64');

// 智能匹配文件夹名（Junk / Junk Email / 垃圾邮件等）
function findBoxName(boxes,want){
  const flat=[];
  (function walk(obj,pfx=''){
    Object.keys(obj||{}).forEach(name=>{
      const box=obj[name]; const path=pfx?`${pfx}${box.delimiter}${name}`:name;
      flat.push(path);
      if(box.children) walk(box.children,path);
    });
  })(boxes);
  const wantLC=String(want||'INBOX').toLowerCase();
  let hit=flat.find(n=>n.toLowerCase()===wantLC);
  if(hit) return hit;
  if(wantLC==='junk'){
    hit=flat.find(n=>/^(junk|junk[-\s]?email|垃圾邮件|垃圾)$/i.test(n));
    if(hit) return hit;
  }
  if(wantLC==='inbox'){
    hit=flat.find(n=>/^inbox$/i.test(n))||'INBOX';
    return hit;
  }
  return flat.find(n=>/^inbox$/i.test(n))||'INBOX';
}

module.exports = async (req,res)=>{
  try{
    const q = req.method==='GET'?req.query:req.body;
    const {
      refresh_token, client_id, email,
      mailbox='INBOX',
      response_type='json',
      debug='0', raw='0'
    } = q;

    if(!refresh_token || !client_id || !email){
      return res.status(400).json({error:'Missing required parameters: refresh_token, client_id, email'});
    }
    const wantDebug = String(debug)==='1';
    const wantRaw   = String(raw)==='1';

    // 1) 换取 IMAP 的 access_token
    const access_token = await get_access_token(refresh_token, client_id);
    if(wantDebug) console.log('[imap] token ok');

    const imap = new Imap({
      user: email,
      xoauth2: genXOAUTH2(email, access_token),
      host: 'outlook.office365.com',
      port: 993,
      tls: true,
      tlsOptions: { rejectUnauthorized: false }
    });

    imap.once('ready', async ()=>{
      try{
        // 2) 找到正确文件夹名并打开（只读）
        const boxes = await new Promise((resolve,reject)=>imap.getBoxes((e,b)=>e?reject(e):resolve(b)));
        const boxName = findBoxName(boxes, mailbox);
        if(wantDebug) console.log('[imap] open box:', boxName);

        await new Promise((resolve,reject)=>imap.openBox(boxName, true, (e)=>e?reject(e):resolve()));

        // 3) 获取最新一封 —— 注意：search 返回 UID！
        const uids = await new Promise((resolve,reject)=>{
          imap.search(['ALL'], (e, results)=>{
            if(e) return reject(e);
            resolve(results.slice(-1)); // 取最后一个 UID（最新）
          });
        });
        if(wantDebug) console.log('[imap] latest UID:', uids);

        if(!uids || !uids.length){
          const html = renderHtmlPage({from:'',subject:'',date:'',htmlBody:'',textBody:'（此目录暂无邮件）'});
          return String(response_type).toLowerCase()==='html'
            ? res.status(200).type('text/html').send(html)
            : res.status(200).json({message:'no messages'});
        }

        // 4) 抓整封原文：一定要加 { uid: true } ！！！
        const f = imap.fetch(uids, { bodies: '', struct: true, uid: true });

        f.on('message',(msg)=>{
          let raw='';

          msg.on('body', (stream)=>{
            stream.on('data', chunk=>{ raw += chunk.toString('utf8'); });
            stream.once('end', async ()=>{
              try{
                if(wantRaw){
                  // 返回源码片段用于排错
                  return res.type('text/plain').send(raw.slice(0,60000));
                }

                const mail = await simpleParser(raw); // 等整封读完再解析
                const data = {
                  send:   mail?.from?.text || '',
                  subject:mail.subject || '',
                  text:   mail.text || '',
                  html:   mail.html || '',
                  date:   mail.date || ''
                };

                if(String(response_type).toLowerCase()==='html'){
                  const page = renderHtmlPage({
                    from: data.send, subject: data.subject, date: data.date,
                    htmlBody: data.html, textBody: data.text
                  });
                  res.status(200).type('text/html').send(page);
                }else{
                  res.status(200).json(data);
                }
              }catch(e){
                console.error('[mailparser] error:', e);
                res.status(500).json({error:String(e.message||e)});
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

    imap.once('error', (err)=>{
      console.error('[imap] error:', err);
      res.status(500).json({error:String(err.message||err)});
    });

    imap.connect();

  }catch(e){
    console.error('[handler] error:', e);
    res.status(500).json({error:String(e.message||e)});
  }
};
