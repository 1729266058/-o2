// Helpers to read Graph message body and MIME raw ($value).
import fetch from 'node-fetch';

const tenantId     = process.env.MS_TENANT_ID || 'common';
const clientId     = process.env.MS_CLIENT_ID;
const clientSecret = process.env.MS_CLIENT_SECRET;
const refreshToken = process.env.MS_REFRESH_TOKEN;
const userId       = process.env.MS_USER_ID || 'me';

function assertEnv() {
  const missing = [];
  if (!clientId) missing.push('MS_CLIENT_ID');
  if (!clientSecret) missing.push('MS_CLIENT_SECRET');
  if (!refreshToken) missing.push('MS_REFRESH_TOKEN');
  if (missing.length) {
    throw new Error('Missing env: ' + missing.join(', '));
  }
}

export async function getAccessToken() {
  assertEnv();
  const form = new URLSearchParams();
  form.set('client_id', clientId);
  form.set('client_secret', clientSecret);
  form.set('refresh_token', refreshToken);
  form.set('grant_type', 'refresh_token');
  form.set('scope', 'https://graph.microsoft.com/.default offline_access');

  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const r = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: form
  });
  if (!r.ok) {
    const t = await r.text();
    throw new Error(`token error ${r.status}: ${t}`);
  }
  const j = await r.json();
  if (!j.access_token) throw new Error('no access_token in token response');
  return j.access_token;
}

function baseUrl() {
  return (userId === 'me')
    ? 'https://graph.microsoft.com/v1.0/me'
    : `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}`;
}

export async function getMessageBody(messageId, accessToken) {
  const r = await fetch(`${baseUrl()}/messages/${messageId}?$select=body,hasAttachments`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  if (!r.ok) return null;
  const j = await r.json();
  return j?.body || null;
}

export async function getMessageMime(messageId, accessToken) {
  const r = await fetch(`${baseUrl()}/messages/${messageId}/$value`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  if (!r.ok) {
    const t = await r.text();
    throw new Error(`mime fetch error ${r.status}: ${t}`);
  }
  return await r.text();
}