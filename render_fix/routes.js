// Express route you can mount into your existing server.
import { getAccessToken, getMessageBody, getMessageMime } from './util-graph.js';
import { parseMailToHtml, wrapHtmlDocument, escapeHtml } from './util-render.js';

export function mountMailHtmlRoutes(app, basePath = '/api') {
  // JSON format
  app.get(`${basePath}/message/:id`, async (req, res) => {
    try {
      const id = req.params.id;
      const token = await getAccessToken();

      const body = await getMessageBody(id, token);
      if (body && body.content && (body.contentType||'text').toLowerCase() === 'html') {
        return res.json({ html: body.content, from: 'graph-body-html' });
      }

      const raw = await getMessageMime(id, token);
      const parsed = await parseMailToHtml(raw);
      if (parsed.html) return res.json({ html: parsed.html, from: parsed.from });

      const snippet = (raw || '').slice(0, 20000);
      return res.json({ html: `<pre>${escapeHtml(snippet)}</pre>`, from: 'raw' });
    } catch (e) {
      console.error(e);
      res.status(500).json({ error: String(e) });
    }
  });

  // Direct HTML render
  app.get(`${basePath}/message/:id/html`, async (req, res) => {
    try {
      const id = req.params.id;
      const token = await getAccessToken();

      let html = null;
      let origin = 'unknown';

      const body = await getMessageBody(id, token);
      if (body && body.content && (body.contentType||'text').toLowerCase() === 'html') {
        html = body.content;
        origin = 'graph-body-html';
      } else {
        const raw = await getMessageMime(id, token);
        const parsed = await parseMailToHtml(raw);
        html = parsed.html;
        origin = parsed.from || 'mime';
      }

      if (!html) html = '<pre>Unable to render message.</pre>';
      const doc = wrapHtmlDocument(html, { title: `Message ${id}`, origin });
      res.type('text/html').send(doc);
    } catch (e) {
      console.error(e);
      res.status(500).type('text/html').send(`<pre>${escapeHtml(String(e))}</pre>`);
    }
  });
}