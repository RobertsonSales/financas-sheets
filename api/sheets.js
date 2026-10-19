// Vercel Serverless Function – proxy seguro entre o front-end e o Google Apps Script.
// O token do Apps Script nunca chega ao navegador.
export default async function handler(req, res) {
  const url = process.env.APPS_SCRIPT_URL;
  const token = process.env.API_TOKEN;
  const accessKey = process.env.ACCESS_KEY; // opcional: senha do operador

  if (!url || !token) {
    return res.status(500).json({ ok: false, error: 'Variáveis APPS_SCRIPT_URL / API_TOKEN não configuradas na Vercel.' });
  }
  if (accessKey && req.headers['x-access-key'] !== accessKey) {
    return res.status(401).json({ ok: false, error: 'Chave de acesso inválida.' });
  }

  try {
    let r;
    if (req.method === 'GET') {
      r = await fetch(`${url}?token=${encodeURIComponent(token)}`, { redirect: 'follow' });
    } else if (req.method === 'POST') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
      r = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ ...body, token }),
        redirect: 'follow',
      });
    } else {
      res.setHeader('Allow', 'GET, POST');
      return res.status(405).json({ ok: false, error: 'Método não permitido' });
    }

    const text = await r.text();
    try {
      const data = JSON.parse(text);
      return res.status(data.ok === false && data.error === 'Não autorizado' ? 401 : 200).json(data);
    } catch {
      console.error('Resposta inválida do Apps Script', r.status, text.slice(0, 500));
      return res.status(502).json({ ok: false, error: `Resposta inválida do Apps Script [${r.status}]` });
    }
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: e.message });
  }
}
