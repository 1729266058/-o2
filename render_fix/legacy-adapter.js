// Legacy adapter: /api/mail-new?message_id=...&response_type=html|json
import { getAccessToken, getMessageBody, getMessageMime } from './util-graph.js';
import { parseMailToHtml, wrapHtmlDocument, escapeHtml } from './util-render.js';

/**
 * Attach legacy endpoint to an existing express app.
 */
export function mountLegacyMailNew(app) {
  app.get('/api/mail-new', async (req, res) => {
    try {
      const id = req.query.message_id || req.query.id;
      const responseType = (req.query.response_type || 'json').toLowerCase();
      if (!id) {
        return res.status(400).json({ error: 'missing message_id' });
      }

      const token = await getAccessToken();

      // Prefer Graph body html
      let html = null, from = 'unknown';
      const body = await getMessageBody(id, token);
      if (body && body.content && (body.contentType||'text').toLowerCase() === 'html') {
        html = body.content; from = 'graph-body-html';
      } else {
        const raw = await getMessageMime(id, token);
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