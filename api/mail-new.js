/**
 * mail-new.js —— 最终版（覆盖此文件即可）
 * 功能/修复：
 * 1) Graph 查询加入 $select=body，确保拿到 body.content
 * 2) Graph 若无 HTML，自动用 /messages/{id}/$value 拉 MIME，并用 mailparser 解析出真正 HTML
 * 3) response_type=html 返回可直接渲染的页面；json 返回单封结构化 JSON
 * 4) IMAP 分支统一优先 html，退回 text
 */

const Imap = require('node-imap');
const { simpleParser } = require('mailparser');
const fetch = require('node-fetch'); // Node 18 也可用全局 fetch；保留这行兼容更好

// =========== 小工具 ===========
function escapeHtml(s = '') {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}
function renderHtmlPage({ from, subject, date, htmlBody, textBody }) {
  const htmlOrText = htmlBody
    ? htmlBody
    : `<pre style="white-space:pre-wrap;">${escapeHtml(textBody || '')}</pre>`;

  return `<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>邮件信息</title>
<style>
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Noto Sans,Arial;background:#f7f7f7;margin:0}
  .wrap{max-width:960px;margin:24px auto;background:#fff;border-radius:12px;box-shadow:0 10px 24px rgba(0,0,0,.06);overflow:hidden}
  .hdr{padding:18px 20px;border-bottom:1px solid #eee}
  .hdr h1{margin:0;font-size:20px}
  .meta{padding:10px 20px;border-bottom:1px solid #f5f5f5;color:#555;font-size:14px}
  .content{padding:16px 20px;min-height:300px}
  .label{font-weight:600;color:#333}
  .warn{padding:12px;background:#fff3cd;border:1px solid #ffeeba;border-radius:8px;color:#856404}
</style>
</head>
<body>
  <div class="wrap">
    <div class="hdr"><h1>邮件信息</h1></div>
    <div class="meta">
      <div><span class="label">发件人：</span>${escapeHtml(from || '')}</div>
      <div><span class="label">主题：</span>${escapeHtml(subject || '')}</div>
      <div><span class="label">日期：</span>${escapeHtml(date || '')}</div>
      <div style="margin-top:8px" class="label">内容：</div>
    </div>
    <div class="content">
      ${htmlOrText || '<div class="warn">此邮件没有可显示的正文或仅包含附件。</div>'}
    </div>
  </div>
</body>
</html>`;
}

// =========== OAuth（保留你原逻辑）===========
async function get_access_token(refresh_token, client_id) {
  const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id,
      grant_type: 'refresh_token',
      refresh_token
    }).toString()
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`HTTP error! status: ${response.status}, response: ${errorText}`);
  }

  const text = await response.text();
  try {
    const data = JSON.parse(text);
    return data.access_token;
  } catch (e) {
    throw new Error(`Failed to parse JSON: ${e.message}, response: ${text}`);
  }
}

const generateAuthString = (user, accessToken) => {
  const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
  return Buffer.from(authString).toString('base64');
};

async function graph_api(refresh_token, client_id) {
  const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id,
      grant_type: 'refresh_token',
      refresh_token,
      scope: 'https://graph.microsoft.com/.default'
    }).toString()
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`HTTP error! status: ${response.status}, response: ${errorText}`);
  }

  const text = await response.text();
  try {
    const data = JSON.parse(text);
    if (data.scope && data.scope.indexOf('https://graph.microsoft.com/Mail.ReadWrite') !== -1) {
      return { access_token: data.access_token, status: true };
    }
    return { access_token: data.access_token, status: false };
  } catch (e) {
    throw new Error(`Failed to parse JSON: ${e.message}, response: ${text}`);
  }
}

// =========== Graph 辅助 ===========
function baseUrl(email) {
  return email
    ? `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(email)}`
    : `https://graph.microsoft.com/v1.0/me`;
}

/** 从 Graph 拉 MIME 源，再解析出 HTML（兜底） */
async function getHtmlFromMime(access_token, email, messageId) {
  try {
    const url = `${baseUrl(email)}/messages/${messageId}/$value`;
    const r = await fetch(url, { headers: { Authorization: `Bearer ${access_token}` } });
    if (!r.ok) {
      const t = await r.text();
      console.error('mime fetch error:', t);
      return null;
    }
    const raw = await r.text();
    const mail = await simpleParser(raw);
    if (mail.html) return mail.html;
    if (mail.text) return `<pre style="white-space:pre-wrap;">${escapeHtml(mail.text)}</pre>`;
    return null;
  } catch (e) {
    console.error('mime parse error:', e);
    return null;
  }
}

/** 读取最新一封（确保 $select=body 等字段） */
async function get_emails(access_token, mailbox) {
  if (!access_token) return [];

  try {
    const url =
      `https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(mailbox)}/messages`
      + `?$top=1&$orderby=receivedDateTime desc`
      + `&$select=id,subject,from,body,bodyPreview,createdDateTime,receivedDateTime`;

    const r = await fetch(url, {
      method: 'GET',
      headers: { Authorization: `Bearer ${access_token}` }
    });

    if (!r.ok) {
      const t = await r.text();
      console.error('graph list error:', t);
      return [];
    }

    const j = await r.json();
    const emails = Array.isArray(j.value) ? j.value : [];

    return emails.map(item => {
      const ct = (item?.body?.contentType || '').toLowerCase();
      return {
        id: item.id,
        send: item?.from?.emailAddress?.address || item?.from?.emailAddress?.name || '',
        subject: item.subject || '',
        text: item.bodyPreview || '',
        html: ct === 'html' ? (item?.body?.content || '') : '',
        date: item.receivedDateTime || item.createdDateTime || '',
      };
    });
  } catch (err) {
    console.error('Error fetching emails:', err);
    return [];
  }
}

// =========== 主处理 ===========
module.exports = async (req, res) => {
  try {
    // 可选密码校验
    const { password } = req.method === 'GET' ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;
    if (password !== expectedPassword && expectedPassword) {
      return res.status(401).json({ error: 'Authentication failed.' });
    }

    // 参数
    const params = req.method === 'GET' ? req.query : req.body;
    let { refresh_token, client_id, email, mailbox, response_type = 'json' } = params;

    if (!refresh_token || !client_id || !email || !mailbox) {
      return res.status(400).json({ error: 'Missing required parameters: refresh_token, client_id, email, or mailbox' });
    }

    console.log("[mail-new] 判断是否 graph_api");
    const graph_api_result = await graph_api(refresh_token, client_id);

    // ===== Graph 分支 =====
    if (graph_api_result.status) {
      console.log("[mail-new] 是 graph_api");

      // 兼容邮箱名映射
      if (mailbox !== "INBOX" && mailbox !== "Junk") mailbox = "inbox";
      if (mailbox === 'INBOX') mailbox = 'inbox';
      if (mailbox === 'Junk')  mailbox = 'junkemail';

      const list = await get_emails(graph_api_result.access_token, mailbox);
      const item = Array.isArray(list) ? list[0] : list;

      if (!item) {
        if (String(response_type).toLowerCase() === 'html') {
          return res.status(200).type('text/html')
                   .send(renderHtmlPage({ from:'', subject:'', date:'', htmlBody:'', textBody:'（此目录暂无邮件）' }));
        }
        return res.status(200).json([]);
      }

      // 若无 HTML，走 MIME 兜底；若仍无，则退回 text
      let htmlBody = item.html || '';
      let textBody = item.text || '';
      if (!htmlBody && item.id) {
        const mimeHtml = await getHtmlFromMime(graph_api_result.access_token, email, item.id);
        if (mimeHtml) htmlBody = mimeHtml;
      }

      if (String(response_type).toLowerCase() === 'html') {
        return res.status(200).type('text/html')
                 .send(renderHtmlPage({
                   from: item.send, subject: item.subject, date: item.date,
                   htmlBody, textBody
                 }));
      }

      return res.status(200).json({
        id: item.id, send: item.send, subject: item.subject,
        text: textBody, html: htmlBody, date: item.date
      });
    }

    // ===== IMAP 分支（保持你的原逻辑并统一输出）=====
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
        // 打开指定邮箱（只读）
        await new Promise((resolve, reject) => {
          imap.openBox(mailbox, true, (err) => err ? reject(err) : resolve());
        });

        const results = await new Promise((resolve, reject) => {
          imap.search(["ALL"], (err, results) => {
            if (err) return reject(err);
            resolve(results.slice(-1)); // 最新一封
          });
        });

        const f = imap.fetch(results, { bodies: "" });
        f.on("message", (msg) => {
          msg.on("body", (stream) => {
            simpleParser(stream, (err, mail) => {
              if (err) throw err;
              const data = {
                send: mail?.from?.text || '',
                subject: mail.subject || '',
                text: mail.text || '',
                html: mail.html || '',
                date: mail.date || ''
              };
              if (String(response_type).toLowerCase() === 'html') {
                return res.status(200).type('text/html')
                         .send(renderHtmlPage({
                           from: data.send, subject: data.subject, date: data.date,
                           htmlBody: data.html, textBody: data.text
                         }));
              }
              return res.status(200).json(data);
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
    console.error('[mail-new] Error:', error);
    res.status(500).json({ error: error.message });
  }
};
