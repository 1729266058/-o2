/**
 * 修复版 mail-new.js —— 直接覆盖本文件即可生效
 * 修复点：
 * 1) Graph 分支支持 response_type=html，返回可直出的 HTML 页面（不再空白）
 * 2) Graph 没有 HTML 时，自动用纯文本 <pre> 兜底（可选开启 MIME 兜底）
 * 3) IMAP 分支统一优先 html、否则退文本
 */

const Imap = require('node-imap');
const { simpleParser } = require('mailparser');
const fetch = require('node-fetch'); // 确保已安装：npm i node-fetch

// ====== 小工具 ======
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

// ====== OAuth 获取 access_token（你原逻辑保留）======
async function get_access_token(refresh_token, client_id) {
  const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      'client_id': client_id,
      'grant_type': 'refresh_token',
      'refresh_token': refresh_token
    }).toString()
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`HTTP error! status: ${response.status}, response: ${errorText}`);
  }
  const responseText = await response.text();
  try {
    const data = JSON.parse(responseText);
    return data.access_token;
  } catch (parseError) {
    throw new Error(`Failed to parse JSON: ${parseError.message}, response: ${responseText}`);
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
      'client_id': client_id,
      'grant_type': 'refresh_token',
      'refresh_token': refresh_token,
      'scope': 'https://graph.microsoft.com/.default'
    }).toString()
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`HTTP error! status: ${response.status}, response: ${errorText}`);
  }

  const responseText = await response.text();

  try {
    const data = JSON.parse(responseText);
    if (data.scope && data.scope.indexOf('https://graph.microsoft.com/Mail.ReadWrite') !== -1) {
      return { access_token: data.access_token, status: true };
    }
    return { access_token: data.access_token, status: false };
  } catch (parseError) {
    throw new Error(`Failed to parse JSON: ${parseError.message}, response: ${responseText}`);
  }
}

// Graph 基础 URL（支持 /me 或 /users/{email}）—— 如果要做 MIME 兜底可用
function baseUrl(email) {
  return email ? `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(email)}`
               : `https://graph.microsoft.com/v1.0/me`;
}

// ====== 读取最新一封（Graph）======
async function get_emails(access_token, mailbox) {
  if (!access_token) {
    console.log("Failed to obtain access token");
    return [];
  }

  try {
    // 增加 $select 以便拿到 body 的 contentType/content 与 id
    const url = `https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(mailbox)}/messages` +
      `?$top=1&$orderby=receivedDateTime desc` +
      `&$select=id,subject,from,body,bodyPreview,createdDateTime,receivedDateTime`;

    const response = await fetch(url, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Authorization': `Bearer ${access_token}`
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error('graph list error:', errorText);
      return [];
    }

    const responseData = await response.json();
    const emails = Array.isArray(responseData.value) ? responseData.value : [];

    const response_emails = emails.map(item => {
      const isHtml = String(item?.body?.contentType || '').toLowerCase() === 'html';
      return {
        id: item.id,
        send: item?.from?.emailAddress?.address || item?.from?.emailAddress?.name || '',
        subject: item.subject || '',
        text: item.bodyPreview || '',
        html: isHtml ? (item?.body?.content || '') : '',
        date: item.receivedDateTime || item.createdDateTime || '',
      };
    });

    return response_emails;
  } catch (error) {
    console.error('Error fetching emails:', error);
    return [];
  }
}

/** （可选）MIME 兜底：当 html 为空时，去拉 $value 再解析出真正的 HTML
async function getHtmlFromMime(access_token, email, messageId) {
  try {
    const url = `${baseUrl(email)}/messages/${messageId}/$value`;
    const r = await fetch(url, { headers: { Authorization: `Bearer ${access_token}` } });
    if (!r.ok) {
      console.error('mime fetch error', await r.text());
      return null;
    }
    const raw = await r.text();
    const mail = await simpleParser(raw);
    return mail.html || (mail.text ? `<pre style="white-space:pre-wrap;">${escapeHtml(mail.text)}</pre>` : null);
  } catch (e) {
    console.error('mime parse error', e);
    return null;
  }
}
*/

// ====== 路由处理 ======
module.exports = async (req, res) => {
  try {
    // 密码校验（可选）
    const { password } = req.method === 'GET' ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;
    if (password !== expectedPassword && expectedPassword) {
      return res.status(401).json({
        error: 'Authentication failed. Please provide valid credentials or contact administrator for access.'
      });
    }

    // 参数
    const params = req.method === 'GET' ? req.query : req.body;
    let { refresh_token, client_id, email, mailbox, response_type = 'json' } = params;

    if (!refresh_token || !client_id || !email || !mailbox) {
      return res.status(400).json({ error: 'Missing required parameters: refresh_token, client_id, email, or mailbox' });
    }

    console.log("判断是否graph_api");
    const graph_api_result = await graph_api(refresh_token, client_id);

    // ==== Graph 分支 ====
    if (graph_api_result.status) {
      console.log("是graph_api");

      // 兼容你原有的邮箱名映射
      if (mailbox !== "INBOX" && mailbox !== "Junk") mailbox = "inbox";
      if (mailbox === 'INBOX') mailbox = 'inbox';
      if (mailbox === 'Junk')  mailbox = 'junkemail';

      const list = await get_emails(graph_api_result.access_token, mailbox);
      const item = Array.isArray(list) ? list[0] : list; // 只取最新一封
      if (!item) {
        if (String(response_type).toLowerCase() === 'html') {
          return res.status(200).type('text/html').send(
            renderHtmlPage({ from: '', subject: '', date: '', htmlBody: '', textBody: '（此目录暂无邮件）' })
          );
        }
        return res.status(200).json([]);
      }

      // 若没有 HTML，先用纯文本兜底；（需要更强兜底时，开启上面的 MIME 方法）
      let htmlBody = item.html || '';
      let textBody = item.text || '';

      /** // 如果你要 MIME 兜底，把下面三行取消注释：
      if (!htmlBody && item.id) {
        const mimeHtml = await getHtmlFromMime(graph_api_result.access_token, email, item.id);
        if (mimeHtml) { htmlBody = mimeHtml; }
      }
      */

      if (String(response_type).toLowerCase() === 'html') {
        const page = renderHtmlPage({
          from: item.send,
          subject: item.subject,
          date: item.date,
          htmlBody,
          textBody
        });
        return res.status(200).type('text/html').send(page);
      } else {
        // JSON（只返回一封，结构更稳）
        return res.status(200).json({
          id: item.id,
          send: item.send,
          subject: item.subject,
          text: textBody,
          html: htmlBody,
          date: item.date
        });
      }
    }

    // ==== IMAP 分支 ====
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
          imap.openBox(mailbox, true, (err, box) => {
            if (err) return reject(err);
            resolve(box);
          });
        });

        const results = await new Promise((resolve, reject) => {
          imap.search(["ALL"], (err, results) => {
            if (err) return reject(err);
            const latestMail = results.slice(-1); // 最新一封
            resolve(latestMail);
          });
        });

        const f = imap.fetch(results, { bodies: "" });

        f.on("message", (msg) => {
          msg.on("body", (stream) => {
            simpleParser(stream, (err, mail) => {
              if (err) throw err;

              const responseData = {
                send: mail?.from?.text || '',
                subject: mail.subject || '',
                text: mail.text || '',
                html: mail.html || '',
                date: mail.date || ''
              };

              if (String(response_type).toLowerCase() === 'json') {
                res.status(200).json(responseData);
              } else if (String(response_type).toLowerCase() === 'html') {
                const page = renderHtmlPage({
                  from: responseData.send,
                  subject: responseData.subject,
                  date: responseData.date,
                  htmlBody: responseData.html,  // 优先 html
                  textBody: responseData.text   // 退文本
                });
                res.status(200).type('text/html').send(page);
              } else {
                res.status(400).json({ error: 'Invalid response_type. Use "json" or "html".' });
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
