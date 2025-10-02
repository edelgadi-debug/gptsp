const express = require('express');
const axios = require('axios');
const mammoth = require('mammoth');
const pdfParse = require('pdf-parse');

// ==== CONFIG ====
const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  SITE_ID,
  DRIVE_ID,
  API_KEY,          // clave propia para proteger el proxy
  PORT = 3000
} = process.env;

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !SITE_ID || !DRIVE_ID) {
  console.error('Faltan variables de entorno obligatorias: TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_ID, DRIVE_ID');
  process.exit(1);
}

// ==== TOKEN CACHE ====
let cachedToken = null;
let tokenExp = 0; // epoch seconds

async function getAccessToken() {
  const now = Math.floor(Date.now() / 1000);
  if (cachedToken && tokenExp - 300 > now) return cachedToken; // renueva 5 min antes de expirar
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    grant_type: 'client_credentials',
    scope: 'https://graph.microsoft.com/.default',
  });
  const { data } = await axios.post(url, body, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
  });
  cachedToken = data.access_token;
  tokenExp = now + (data.expires_in || 3600);
  return cachedToken;
}

function requireApiKey(req, res, next) {
  if (!API_KEY) return next(); // sin clave durante pruebas locales
  const key = req.header('x-api-key');
  if (key === API_KEY) return next();
  res.status(401).json({ error: 'Unauthorized' });
}

async function graphGet(path, params) {
  const token = await getAccessToken();
  const url = `https://graph.microsoft.com/v1.0${path}`;
  const { data } = await axios.get(url, {
    headers: { Authorization: `Bearer ${token}` },
    params
  });
  return data;
}

async function listChildrenByPath(path, params) {
  const enc = encodeURI(path);
  return await graphGet(`/sites/${SITE_ID}/drives/${DRIVE_ID}/root:/${enc}:/children`, params);
}

async function listAllFilesUnderPath(pathPrefix) {
  const results = [];
  async function walk(path) {
    const page = await listChildrenByPath(path, { $top: 200, $select: 'id,name,webUrl,lastModifiedDateTime,parentReference,file,folder' });
    for (const it of (page.value || [])) {
      if (it.folder) {
        await walk(`${path}/${it.name}`);
      } else if (it.file) {
        results.push(it);
      }
    }
    // (Opcional) Si tuvieras muchas páginas: seguir @odata.nextLink (no suele hacer falta en carpetas pequeñas)
  }
  await walk(pathPrefix);
  return results;
}

function normalizeForMatch(s) {
  return (s || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, ''); // elimina tildes
}
function scoreChunk(chunk, query) {
  const terms = normalizeForMatch(query).split(/\s+/).filter(Boolean);
  const lc = normalizeForMatch(chunk);
  return terms.reduce((acc, t) => acc + (lc.split(t).length - 1), 0);
}

async function downloadContentById(itemId) {
  const token = await getAccessToken();
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drives/${DRIVE_ID}/items/${itemId}/content`;
  const resp = await axios.get(url, {
    headers: { Authorization: `Bearer ${token}` },
    responseType: 'arraybuffer',
    maxRedirects: 5
  });
  return Buffer.from(resp.data);
}

function extOf(name = '') {
  const m = name.toLowerCase().match(/\.([a-z0-9]+)$/);
  return m ? m[1] : '';
}

async function extractTextFromBuffer(buf, name) {
  const ext = extOf(name);
  if (ext === 'pdf') {
    const result = await pdfParse(buf);
    return result.text || '';
  }
  if (ext === 'docx') {
    const result = await mammoth.extractRawText({ buffer: buf });
    return result.value || '';
  }
  if (ext === 'txt') {
    return buf.toString('utf8');
  }
  // Podrías agregar más: 'pptx', 'html', etc.
  return ''; // tipo no soportado -> sin texto
}

function chunkText(txt, maxChars = 1200) {
  const chunks = [];
  let i = 0;
  while (i < txt.length && chunks.length < 50) { // cap de seguridad
    chunks.push(txt.slice(i, i + maxChars));
    i += maxChars;
  }
  return chunks;
}

function scoreChunk(chunk, query) {
  if (!query) return 0;
  const terms = query.toLowerCase().split(/\s+/).filter(Boolean);
  const lc = chunk.toLowerCase();
  return terms.reduce((acc, t) => acc + (lc.split(t).length - 1), 0);
}

// ==== APP ====
const app = express();
app.use(express.json({ limit: '2mb' }));
app.use(requireApiKey);

// Health
app.get('/', (req, res) => res.json({ ok: true, service: 'sp-knowledge-proxy' }));

// Listar raíz
app.get('/root', async (req, res) => {
  try {
    const data = await graphGet(`/sites/${SITE_ID}/drives/${DRIVE_ID}/root/children`, {
      $top: req.query.$top,
      $skip: req.query.$skip,
      $select: req.query.$select,
      $expand: req.query.$expand,
    });
    res.json(data);
  } catch (e) {
    console.error(e?.response?.data || e.message);
    res.status(e?.response?.status || 500).json(e?.response?.data || { error: e.message });
  }
});

// Listar carpeta por ruta
app.get('/folder', async (req, res) => {
  try {
    const path = req.query.path;
    if (!path) return res.status(400).json({ error: 'Missing query param: path' });
    const enc = encodeURI(path);
    const data = await graphGet(`/sites/${SITE_ID}/drives/${DRIVE_ID}/root:/${enc}:/children`);
    res.json(data);
  } catch (e) {
    console.error(e?.response?.data || e.message);
    res.status(e?.response?.status || 500).json(e?.response?.data || { error: e.message });
  }
});

// Descargar por ruta o id
app.get('/download', async (req, res) => {
  try {
    const { path, id } = req.query;
    if (!path && !id) return res.status(400).json({ error: 'Provide ?path= or ?id=' });
    const token = await getAccessToken();
    const url = path
      ? `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drives/${DRIVE_ID}/root:/${encodeURI(path)}:/content`
      : `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drives/${DRIVE_ID}/items/${id}/content`;
    const resp = await axios.get(url, {
      headers: { Authorization: `Bearer ${token}` },
      responseType: 'stream',
      maxRedirects: 5,
    });
    if (resp.headers['content-type']) res.setHeader('Content-Type', resp.headers['content-type']);
    if (resp.headers['content-disposition']) res.setHeader('Content-Disposition', resp.headers['content-disposition']);
    resp.data.pipe(res);
  } catch (e) {
    console.error(e?.response?.data || e.message);
    res.status(e?.response?.status || 500).json(e?.response?.data || { error: e.message });
  }
});

// ==== RETRIEVE: busca y devuelve snippets ====
app.post('/retrieve', async (req, res) => {
  try {
    const {
      query,
      pathPrefix,             // ej: "General/Desarrollo_organizacional/Manuales"
      topK = 6,
      maxCharsPerChunk = 1200,
      fileTypes = ['pdf', 'docx', 'txt'],
      includeFileText = false
    } = req.body || {};

    if (!query || typeof query !== 'string') {
      return res.status(400).json({ error: 'query (string) es obligatorio' });
    }

    // 1) Obtener candidatos: por carpeta (si hay pathPrefix) o por búsqueda en Graph
    let items = [];
    if (pathPrefix && pathPrefix.trim() !== '') {
      // Escaneo directo de la carpeta (recursivo)
      items = await listAllFilesUnderPath(pathPrefix);
    } else {
      // Búsqueda global en el drive con Graph Search
      const searchPath = `/sites/${SITE_ID}/drives/${DRIVE_ID}/root/search(q='${encodeURIComponent(query)}')`;
      const searchRes = await graphGet(searchPath, { $top: Math.max(topK * 4, 20) });
      items = (searchRes.value || []).filter(it => it.file);
    }

    // 2) Filtrar por tipos de archivo soportados
    items = items.filter(it => fileTypes.includes(extOf(it.name)));

    // 3) Procesar un subconjunto razonable (evita cargas enormes)
    const filesToProcess = items.slice(0, Math.max(topK * 4, 20));

    // 4) Descargar, extraer texto, trocear y puntuar
    const snippets = [];
    for (const it of filesToProcess) {
      const buf = await downloadContentById(it.id);
      const text = await extractTextFromBuffer(buf, it.name);
      if (!text) continue;

      const chunks = chunkText(text, maxCharsPerChunk);
      const scored = chunks.map(c => ({
        text: c,
        score: scoreChunk(c, query),
        file: {
          itemId: it.id,
          name: it.name,
          path: `${it.parentReference?.path?.replace('/drive', '') || ''}/${it.name}`,
          webUrl: it.webUrl,
          contentType: it.file?.mimeType,
          lastModifiedDateTime: it.lastModifiedDateTime
        }
      }));

      scored.sort((a, b) => b.score - a.score);
      if (scored[0]) snippets.push(scored[0]); // mejor chunk de ese archivo
    }

    // 5) Ordenar globalmente y limitar a topK
    snippets.sort((a, b) => (b.score || 0) - (a.score || 0));
    const limited = snippets.slice(0, topK);
    const combinedContext = limited.map(s => s.text).join('\n---\n');

    const resp = {
      query,
      usedParams: { pathPrefix, topK, maxCharsPerChunk, fileTypes, includeFileText },
      snippets: limited,
      topFiles: limited.map(s => s.file),
      combinedContext
    };

    if (includeFileText && limited[0]) {
      resp.fullText = limited[0].text;
    }

    res.json(resp);
  } catch (e) {
    console.error(e?.response?.data || e.message);
    res.status(e?.response?.status || 500).json(e?.response?.data || { error: e.message });
  }
});

app.listen(PORT, () => console.log(`Proxy running on http://localhost:${PORT}`));
