import { simpleParser } from 'mailparser';

export async function parseMailToHtml(raw) {
  if (!raw) return { html: null, from: 'no-raw' };
  const mail = await simpleParser(raw);
  let html = mail.html || null;
  if (!html && mail.text) {
    html = `<pre style="white-space:pre-wrap;">${escapeHtml(mail.text)}</pre>`;
  }
  if (html && mail.attachments && mail.attachments.length) {
    for (const a of mail.attachments) {
      if (a.cid && a.content && a.content.length <= 1500000) {
        const mime = a.contentType || 'application/octet-stream';
        const b64 = Buffer.isBuffer(a.content) ? a.content.toString('base64') : Buffer.from(a.content).toString('base64');
        const dataUri = `data:${mime};base64,${b64}`;
        const cidPattern = new RegExp(`cid:${escapeRegExp(a.cid)}`, 'gi');
        html = html.replace(cidPattern, dataUri);
      }
    }
  }
  return { html, from: mail.html ? 'mime-html' : (mail.text ? 'mime-text' : 'mime-unknown') };
}

export function wrapHtmlDocument(innerHtml, { title = 'Message', origin = '' } = {}) {
  return `<!doctype html>
<html><head>
<meta charset="utf-8"/><meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>${escapeHtml(title)}</title>
<style>
  body{margin:0;padding:0;background:#f7f7f7;}
  .container{max-width:960px;margin:0 auto;padding:16px;}
  .card{background:#fff;border-radius:12px;box-shadow:0 8px 24px rgba(0,0,0,.08);padding:16px;}
  .meta{font:12px/1.4 system-ui, -apple-system, Segoe UI, Roboto, Noto Sans, Arial; color:#666;margin-bottom:8px}
  .content{min-height:200px}
</style>
</head>
<body>
  <div class="container">
    <div class="meta">Rendered via <b>${escapeHtml(origin)}</b></div>
    <div class="card content">${innerHtml}</div>
  </div>
</body></html>`;
}

export function escapeHtml(s='') {
  return s.replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
}
function escapeRegExp(s='') {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}