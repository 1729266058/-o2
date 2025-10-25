/**
 * mail-new.js —— IMAP OAuth2 稳定版 + Graph 兼容 + 大邮件降级
 * - IMAP：按 UID 抓最新一封，先缓冲完整原文再解析；优先 html，无则 textAsHtml；超大 HTML 自动降级
 * - Graph：查询时 $select=body，支持 response_type=html，无 html 时退回 <pre>
 * - 兼容参数：refresh_token, client_id, email, mailbox, response_type=html|json
 * - 诊断辅助：&raw=1（返回前 60KB 源码片段），&debug=1（日志）
 */

const Imap = require('node-imap');
const { simpleParser } = require('mailparser');
const fetch = require('node-fetch');

// 超过这个长度的 HTML（约 1.8MB）用 textAsHtml 降级以避免平台截断造成空白
const MAX_HTML_SIZE = parseInt(process.env.MAX_HTML_SIZE || '1800000', 10);

// ---------- 工具 ----------
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
.content{padding:16px 20px;min-height:300px}.label{font-weight:600;color:#333}
.warn{padding:12px;background:#fff3cd;border:1px solid #ffeeba;border-radius:8px;color:#856404;margin-bottom:12px}
</style></head><body>
<div class="wrap">
  <div class="hdr"><h1>邮件信息</h1></div>
  <div class="meta">
    <div><span class="label">发件人：</span>${escapeHtml(from||'')}</div>
    <div><span class="label">主题：</span>${escapeHtml(subject||'')}</div>
    <div><span class="label">日期：</span>${escapeHtml(String(date||''))}</div>
    <div style="margin-top:8px" class="label">内容：</div>
  </div>
  <div class="content">
    ${notice}
    ${htmlOrText || '<div class="warn">⚠️ 此邮件没有可显示的正文或仅包含附件/图片。</div>'}
  </div>
</div>
</body></html>`;
}

// ---------- OAuth ----------
async function get_access_token(refresh_token, client_id) {
  const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id, grant_type: 'refresh_token', refresh_token
    }).toString()
  });
  if (!response.ok) throw new Error(`HTTP error! ${response.status}: ${await response.text()}`);
  const data = await response.json();
  if (!data.access_token) throw new Error('no access_token');
  return data.access_token;
}

const generateAuthString = (user, accessToken) =>
  Buffer.from(`user=${user}\x01auth=Bearer ${accessToken}\x01\x01`).toString('base64');

async function graph_api(refresh_token, client_id) {
  const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id, grant_type: 'refresh_token', refresh_token, scope: 'https://graph.microsoft.com/.default'
    }).toString()
  });
  if (!response.ok) throw new Error(`HTTP error! ${response.status}: ${await response.text()}`);
  const data = await response.json();
  const ok = data.scope && data.scope.indexOf('https://graph.microsoft.com/Mail.ReadWrite') !== -1;
  return { access_token: data.access_token, status: !!ok };
}

// ---------- Graph ----------
async function get_emails(access_token, mailbox) {
  if (!access_token) return [];
  try {
    const url = `https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(mailbox)}/messages`
      + `?$top=1&$orderby=receivedDateTime desc`
      + `&$select=id,subject,from,body,bodyPreview,createdDateTime,receivedDateTime`;
    const r = await fetch(url, { headers: { Authorization: `Bearer ${access_token}` } });
    if (!r.ok) { console.error('graph list error:', await r.text()); return []; }
    const j = await r.json();
    const arr = Array.isArray(j.value) ? j.value : [];
    return arr.map(it => {
      const ct = (it?.body?.contentType || '').toLowerCase();
      return {
        id: it.id,
        send: it?.from?.emailAddress?.address || it?.from?.emailAddress?.name || '',
        subject: it.subject || '',
        text: it.bodyPreview || '',
        html: ct === 'html' ? (it?.body?.content || '') : '',
        date: it.receivedDateTime || it.createdDateTime || '',
      };
    });
  } catch (e) {
    console.error('graph fetch err:', e);
    return [];
  }
}

// ---------- IMAP 文件夹名智能匹配（Junk/Junk Email/垃圾邮件等） ----------
function findBoxName(boxes, want) {
  const flat = [];
  (function walk(obj, p=''){
    Object.keys(obj||{}).forEach(name=>{
      const box=obj[name]; const path=p?`${p}${box.delimiter}${name}`:name;
      flat.push(path);
      if (box.children) walk(box.children, path);
    });
  })(boxes);
  const w = String(want||'INBOX').toLowerCase();
  let hit = flat.find(n=>n.toLowerCase()===w); if (hit) return hit;
  if (w==='junk'){ hit = flat.find(n=>/^(junk|junk[-\s]?email|垃圾邮件|垃圾)$/i.test(n)); if (hit) return hit; }
  if (w==='inbox'){ hit = flat.find(n=>/^inbox$/i.test(n)) || 'INBOX'; return hit; }
  return flat.find(n=>/^inbox$/i.test(n)) || 'INBOX';
}

// ---------- 主处理 ----------
module.exports = async (req, res) => {
  try {
    // 认证（可选）
    const { password } = req.method === 'GET' ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;
    if (expectedPassword && password !== expectedPassword) {
      return res.status(401).json({ error: 'Authentication failed.' });
    }

    // 参数
    const params = req.method === 'GET' ? req.query : req.body;
    let { refresh_token, client_id, email, mailbox, response_type = 'json', raw = '0', debug = '0' } = params;

    if (!refresh_token || !client_id || !email || !mailbox) {
      return res.status(400).json({ error: 'Missing required parameters: refresh_token, client_id, email, or mailbox' });
    }
    const wantRaw   = String(raw)==='1';
    const wantDebug = String(debug)==='1';

    // 判断是否 Graph 令牌
    let graph_ok = false, graph_token = null;
    try {
      const gr = await graph_api(refresh_token, client_id);
      graph_ok = gr.status; graph_token = gr.access_token;
    } catch { /* ignore */ }

    // ===== Graph 分支 =====
    if (graph_ok) {
      if (mailbox !== "INBOX" && mailbox !== "Junk") mailbox = "inbox";
      if (mailbox === 'INBOX') mailbox = 'inbox';
      if (mailbox === 'Junk')  mailbox = 'junkemail';

      const list = await get_emails(graph_token, mailbox);
      const item = Array.isArray(list) ? list[0] : list;

      if (!item) {
        if (String(response_type).toLowerCase()==='html') {
          return res.status(200).type('text/html')
            .send(renderHtmlPage({ from:'', subject:'', date:'', htmlBody:'', textBody:'（此目录暂无邮件）' }));
        }
        return res.status(200).json([]);
      }

      const htmlBody = item.html || '';
      const textBody = item.text || '';

      if (String(response_type).toLowerCase()==='html') {
        return res.status(200).type('text/html')
          .send(renderHtmlPage({ from: item.send, subject: item.subject, date: item.date, htmlBody, textBody }));
      }
      return res.status(200).json({ id:item.id, send:item.send, subject:item.subject, text:textBody, html:htmlBody, date:item.date });
    }

    // ===== IMAP 分支（仅使用 IMAP OAuth2）=====
    const access_token = await get_access_token(refresh_token, client_id);
    const authString = generateAuthString(email, access_token);

    const imap = new Imap({
      user: email,
      xoauth2: authString,
      host: 'outlook.office365.com',
      port: 993,
      tls: true,
      tlsOptions: { rejectUnauthorized: false }
    });

    imap.once("ready", async () => {
      try {
        const boxes = await new Promise((resolve,reject)=>imap.getBoxes((e,b)=>e?reject(e):resolve(b)));
        const boxName = findBoxName(boxes, mailbox);
        if (wantDebug) console.log('[imap] open box:', boxName);

        await new Promise((resolve,reject)=>imap.openBox(boxName, true, (e)=>e?reject(e):resolve()));

        const uids = await new Promise((resolve,reject)=>{
          imap.search(["ALL"], (err, results) => {
            if (err) return reject(err);
            resolve(results.slice(-1)); // 最新一封的 UID
          });
        });
        if (wantDebug) console.log('[imap] latest UID:', uids);

        if (!uids || !uids.length) {
          const page = renderHtmlPage({ from:'', subject:'', date:'', htmlBody:'', textBody:'（此目录暂无邮件）' });
          return String(response_type).toLowerCase()==='html'
            ? res.status(200).type('text/html').send(page)
            : res.status(200).json({ message:'no messages' });
        }

        // **关键**：按 UID 抓整封；先缓冲完整原文，再解析
        const f = imap.fetch(uids, { bodies: "", struct: true, uid: true });

        f.on("message", (msg) => {
          let rawBuf = '';

          msg.on("body", (stream) => {
            stream.on('data', chunk => { rawBuf += chunk.toString('utf8'); });
            stream.once('end', async () => {
              try {
                if (wantRaw) {
                  return res.status(200).type('text/plain').send(rawBuf.slice(0, 60000));
                }

                const mail = await simpleParser(rawBuf);
                let html = mail.html || '';
                const text = mail.text || '';
                const textAsHtml = mail.textAsHtml || (text ? `<pre style="white-space:pre-wrap;">${escapeHtml(text)}</pre>` : '');

                let hint = '';
                if (html && html.length > MAX_HTML_SIZE) {
                  hint = `⚠️ 正文较大（${(html.length/1024/1024).toFixed(2)} MB），已切换为简化视图以避免平台截断。`;
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

                if (String(response_type).toLowerCase()==='html') {
                  const page = renderHtmlPage({
                    from: data.send, subject: data.subject, date: data.date,
                    htmlBody: data.html, textBody: data.text, hint
                  });
                  res.status(200).type('text/html').send(page);
                } else {
                  res.status(200).json(data);
                }
              } catch (e) {
                res.status(500).json({ error: String(e.message || e) });
              }
            });
          });
        });

        f.once("end", () => imap.end());
      } catch (err) {
        imap.end();
        res.status(500).json({ error: err.message });
      }
    });

    imap.once('error', (err) => {
      console.error('IMAP error:', err);
      res.status(500).json({ error: err.message });
    });

    imap.connect();

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: error.message });
  }
};
