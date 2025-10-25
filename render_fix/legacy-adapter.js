// Legacy adapter: /api/mail-new?message_id=...&response_type=html|json
import { getAccessToken, getMessageBody, getMessageMime } from './util-graph.js';
import { parseMailToHtml, wrapHtmlDocument, escapeHtml } from './util-render.js';

/**
 * Attach legacy endpoint to an existing express app.
 */

// helper: resolve message by Graph id or internetMessageId
async function resolveMessageId(id, token) {
  // If ID looks like an internetMessageId (contains '@' or starts with '<')
  const looksLikeIMI = /@/.test(id) || /^<.*>$/.test(id);
  if (!looksLikeIMI) return id;

  // search by internetMessageId
  const quoted = id.replace(/\\/g,'\\\\').replace(/'/g,"\\'");
  const base = (process.env.MS_USER_ID && process.env.MS_USER_ID !== 'me')
    ? `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(process.env.MS_USER_ID)}`
    : `https://graph.microsoft.com/v1.0/me`;
  const url = `${base}/messages?$filter=internetMessageId eq '${encodeURIComponent(quoted)}'&$select=id,internetMessageId`;
  const fetchMod = await import('node-fetch');
  const fetch = fetchMod.default;
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` }});
  if (r.ok) {
    const j = await r.json();
    if (j && j.value && j.value.length) {
      return j.value[0].id;
    }
  }
  // fallback: return original
  return id;
}

export function mountLegacyMailNew(app) {
  app.get('/api/mail-new', async (req, res) => {
    try {
      const id = req.query.message_id || req.query.id;
      const responseType = (req.query.response_type || 'json').toLowerCase();
      if (!id) {
        return res.status(400).json({ error: 'missing message_id' });
      }

      const token = await getAccessToken();
      const realId = await resolveMessageId(id, token);
      // Prefer Graph body html
      let html = null, from = 'unknown';
      const body = await getMessageBody(realId, token);
      if (body && body.content && (body.contentType||'text').toLowerCase() === 'html') {
        html = body.content; from = 'graph-body-html';
      } else {
        const raw = await getMessageMime(realId, token);
        const parsed = await parseMailToHtml(raw);
        html = parsed.html; from = parsed.from || 'mime';
      }

      // Fallbacks
      if (!html) {
        html = '<pre style="white-space:pre-wrap;">（此邮件没有可显示的正文或仅包含附件）</pre>';
        from = 'no-body';
      }

      if (responseType === 'html') {
        // Return a minimal HTML document (old behavior tends to embed directly)
        const doc = wrapHtmlDocument(html, { title: `Message ${id}`, origin: from });
        return res.type('text/html').send(doc);
      } else {
        return res.json({ html, from });
      }
    } catch (e) {
      console.error(e);
      res.status(500).json({ error: String(e) });
    }
  });
}