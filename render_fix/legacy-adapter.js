// Legacy adapter: /api/mail-new?message_id=...&response_type=html|json
import { getAccessToken, getMessageBody, getMessageMime } from './util-graph.js';
import { parseMailToHtml, wrapHtmlDocument, escapeHtml } from './util-render.js';

// helper: resolve message by Graph id or internetMessageId
async function resolveMessageId(id, token) {
  // 如果像 internetMessageId（包含 @ 或以 < 开头）
  const looksLikeIMI = /@/.test(id) || /^<.*>$/.test(id);
  if (!looksLikeIMI) return id;

  // 按 internetMessageId 查询出真实 Graph id
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
    if (j && j.value && j.value.length) return j.value[0].id;
  }
  return id; // 找不到就用原值
}

/**
 * Attach legacy endpoint to an existing express app.
 */
export function mountLegacyMailNew(app) {
  app.get('/api/mail-new', async (req, res) => {
    try {
      const idParam = req.query.message_id || req.query.id;
      const responseType = (req.query.response_type || 'json').toLowerCase();
      const wantRaw   = String(req.query.raw||'0') === '1';
      const wantDebug = String(req.query.debug||'0') === '1';

      if (!idParam) {
        return res.status(400).json({ error: 'missing message_id' });
      }

      const token = await getAccessToken();
      if (wantDebug) console.log('[mail-new] idParam=', idParam);

      const realId = await resolveMessageId(idParam, token);
      if (wantDebug) console.log('[mail-new] realId=', realId);

      // 需要原始 MIME 直接返回（排错用）
      if (wantRaw) {
        const raw0 = await getMessageMime(realId, token);
        return res.type('text/plain').send(raw0.slice(0, 60000));
      }

      // 1) 尝试直接 Graph body
      const body = await getMessageBody(realId, token);
      if (wantDebug) console.log('[mail-new] graph body contentType=', body?.contentType, 'len=', body?.content?.length);

      let html = null, from = 'unknown';
      if (body && body.content && (body.contentType || 'text').toLowerCase() === 'html') {
        html = body.content; from = 'graph-body-html';
      } else {
        // 2) MIME 兜底（最稳）
        const raw = await getMessageMime(realId, token);
        if (wantDebug) console.log('[mail-new] MIME length=', raw?.length);
        const parsed = await parseMailToHtml(raw);
        html = parsed.html; from = parsed.from || 'mime';
      }

      // 3) 兜底提示（无正文或仅附件）
      if (!html) {
        html = '<div style="padding:12px;background:#fff3cd;border:1px solid #ffeeba;border-radius:8px;color:#856404">'
             + '⚠️ 此邮件没有可显示的正文或仅包含附件。'
             + '你可以在 URL 末尾加 <code>&raw=1</code> 查看原始 MIME 片段，或加 <code>&debug=1</code> 查看服务端日志。'
             + '</div>';
        from = 'no-body';
      }

      if (responseType === 'html') {
        const doc = wrapHtmlDocument(html, { title: `Message ${idParam}`, origin: from });
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
