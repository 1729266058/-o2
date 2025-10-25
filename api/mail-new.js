/**
 * mail-new.js —— IMAP OAuth2 强制版（覆盖此文件即可）
 * 适用：只有 Outlook IMAP/SMTP OAuth2 刷新令牌（不是 Graph Mail.Read）
 * 功能：
 *  - 仅用 IMAP + XOAUTH2 读取“最新一封”并渲染
 *  - 自动匹配邮箱夹（INBOX / Junk / Junk Email / 垃圾邮件 等）
 *  - response_type=html：输出完整可渲染页面；json：输出结构化 JSON
 *  - ?debug=1 打印调试日志；?raw=1 返回前 60KB 源码片段（排错）
 */

const Imap = require('node-imap');
const { simpleParser } = require('mailparser');
const fetch = require('node-fetch'); // 如 Node18 可用全局 fetch; 保留兼容

// ---------- 工具 ----------
function escapeHtml(s=''){return String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function renderHtmlPage({from,subject,date,htmlBody,textBody}){
  const htmlOrText = htmlBody ? htmlBody : `<pre style="white-space:pre-wrap;">${escapeHtml(textBody||'')}</pre>`;
  return `<!doctype html><html><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
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
  <div class="content">${htmlOrText || '<div class="warn">此邮件没有可显示的正文或仅包含附件。</div>'}</div>
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
const generateAuthString=(user,accessToken)=>Buffer.from(`user=${user}\x01auth=Bearer ${accessToken}\x01\x01`).toString('base64');

// 智能匹配邮箱夹名（Junk / Junk Email / 垃圾邮件等）
function findMailboxName(boxes, want){
  const flat = [];
  const walk=(obj,pfx='')=>{
    Object.keys(obj||{}).forEach(name=>{
      const box=obj[name]; const path=pfx?`${pfx}${box.delimiter}${name}`:name;
      flat.push(path);
      if(box.children) walk(box.children,path);
    });
  };
  walk(boxes);
  const wantLC = want.toLowerCase();
  // 直匹配
  let hit = flat.find(n=>n.toLowerCase()===wantLC);
  if(hit) return hit;
  // Junk 特殊
  if(wantLC==='junk'){
    hit = flat.find(n=>/^(junk|junk[-\s]?email|垃圾|垃圾邮件)$/i.test(n));
    if(hit) return hit;
  }
  // INBOX 兼容
  if(wantLC==='inbox'){
    hit = flat.find(n=>/^inbox$/i.test(n)) || 'INBOX';
    return hit;
  }
  // 退回 INBOX
  return flat.find(n=>/^inbox$/i.test(n)) || 'INBOX';
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

    // 1) 换 IMAP 用的 Access Token
    const access_token = await get_access_token(refresh_token, client_id);
    if(wantDebug) console.log('[imap] got access token');

    const authString = generateAuthString(email, access_token);

    const imap = new Imap({
      user: email,
      xoauth2: authString,
      host: 'outlook.office365.com',
      port: 993,
      tls: true,
      tlsOptions: { rejectUnauthorized: false }
    });

    imap.once('ready', async ()=>{
      try{
        // 2) 罗列文件夹并智能定位
        const boxes = await new Promise((resolve,reject)=>{
          imap.getBoxes((err, boxes)=> err?reject(err):resolve(boxes));
        });
        const boxName = findMailboxName(boxes, mailbox);
        if(wantDebug) console.log('[imap] open box:', boxName);

        // 3) 打开文件夹（只读）
        await new Promise((resolve,reject)=>{
          imap.openBox(boxName, true, (err)=> err?reject(err):resolve());
        });

        // 4) 查找最新一封
        const ids = await new Promise((resolve,reject)=>{
          imap.search(['ALL'], (err, results)=>{
            if(err) return reject(err);
            const last = results.slice(-1);
            resolve(last);
          });
        });
        if(wantDebug) console.log('[imap] latest ids:', ids);
        if(!ids || !ids.length){
          const emptyHtml = renderHtmlPage({from:'',subject:'',date:'',htmlBody:'',textBody:'（此目录暂无邮件）'});
          return String(response_type).toLowerCase()==='html'
            ? res.status(200).type('text/html').send(emptyHtml)
            : res.status(200).json({message:'no messages'});
        }

        if(wantRaw){
          // 原始源码片段（便于排错）
          const f = imap.fetch(ids, { bodies: '' });
          f.on('message', (msg)=>{
            msg.on('body', async (stream)=>{
              let rawBuf=''; for await (const chunk of stream) rawBuf+=chunk.toString('utf8');
              return res.type('text/plain').send(rawBuf.slice(0,60000));
            });
          });
          f.once('end', ()=> imap.end());
          return;
        }

        // 5) 取整封并用 mailparser 解析
        const f = imap.fetch(ids, { bodies: '' });
        f.on('message', (msg)=>{
          msg.on('body', async (stream)=>{
            try{
              const mail = await simpleParser(stream);
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
