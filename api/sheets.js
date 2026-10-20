// Vercel Serverless Function – proxy seguro entre o front-end e o Google Apps Script.
// O token do Apps Script nunca chega ao navegador. A sessão (login com senha) vem no header x-session.
export default async function handler(req, res) {
  const url = process.env.APPS_SCRIPT_URL;
  const token = process.env.API_TOKEN;
  const accessKey = process.env.ACCESS_KEY; // opcional e agora dispensável: a senha do app substitui esta chave

  res.setHeader('Cache-Control', 'no-store');

  if (!url || !token) {
    return res.status(500).json({ ok: false, error: 'Variáveis APPS_SCRIPT_URL / API_TOKEN não configuradas na Vercel.' });
  }
  if (accessKey && req.headers['x-access-key'] !== accessKey) {
    return res.status(401).json({ ok: false, error: 'Chave de acesso inválida.' });
  }

  const session = String(req.headers['x-session'] || '');

  try {
    let payload;
    if (req.method === 'GET') {
      payload = { action: 'list' };
    } else if (req.method === 'POST') {
      payload = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
    } else {
      res.setHeader('Allow', 'GET, POST');
      return res.status(405).json({ ok: false, error: 'Método não permitido' });
    }

    // token e session são sempre definidos aqui, nunca pelo corpo enviado pelo navegador
    const r = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ ...payload, token, session }),
      redirect: 'follow',
    });

    const text = await r.text();
    try {
      const data = JSON.parse(text);
      const unauthorized = data.ok === false && (data.error === 'Não autorizado' || data.code === 'AUTH_REQUIRED');
      return res.status(unauthorized ? 401 : 200).json(data);
    } catch {
      console.error('Resposta inválida do Apps Script', r.status, text.slice(0, 500));
      // Mostra um trecho do texto devolvido (sem HTML) para facilitar o diagnóstico
      const hint = text.replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 200);
      return res.status(502).json({ ok: false, error: `Resposta inválida do Apps Script [${r.status}]${hint ? ': ' + hint : ''}` });
    }
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: e.message });
  }
}
