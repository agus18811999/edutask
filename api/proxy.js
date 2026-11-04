/** Vercel Serverless Function. Extra infrastructure file required for reliable GAS access. */
export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  if (req.method !== 'POST') return res.status(405).json({ success: false, code: 'METHOD_NOT_ALLOWED', message: 'POST only.' });
  if (!process.env.EDUTASK_GAS_URL || !/^https:\/\/script\.google\.com\/macros\/s\//.test(process.env.EDUTASK_GAS_URL)) {
    return res.status(503).json({ success: false, code: 'CONFIG_ERROR', message: 'EDUTASK_GAS_URL belum dikonfigurasi pada Vercel.' });
  }
  try {
    const origin = req.headers.origin;
    if (origin && new URL(origin).host !== req.headers.host) {
      return res.status(403).json({ success: false, code: 'FORBIDDEN', message: 'Cross-origin API access denied.' });
    }
    const body = typeof req.body === 'string' ? req.body : JSON.stringify(req.body || {});
    if (Buffer.byteLength(body, 'utf8') > 30000) return res.status(413).json({ success: false, code: 'TOO_LARGE', message: 'Payload terlalu besar.' });
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 26000);
    let upstream;
    try {
      upstream = await fetch(process.env.EDUTASK_GAS_URL, {
        method: 'POST', headers: { 'Content-Type': 'text/plain; charset=utf-8' },
        body, redirect: 'follow', signal: controller.signal, cache: 'no-store'
      });
    } finally { clearTimeout(timer); }
    const data = await upstream.text();
    let json;
    try { json = JSON.parse(data); } catch (_) {
      return res.status(502).json({ success: false, code: 'GAS_RESPONSE_ERROR', message: 'Backend belum terhubung atau respons bukan JSON. Periksa deployment Apps Script.' });
    }
    return res.status(upstream.ok ? 200 : 502).json(json);
  } catch (error) {
    return res.status(502).json({ success: false, code: 'UPSTREAM_UNAVAILABLE', message: 'Tidak dapat menghubungi Google Apps Script. Periksa URL, izin deployment, dan koneksi.' });
  }
}
