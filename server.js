import cors from "cors";
import express from "express";
import { readFile } from "node:fs/promises";
import { randomUUID, randomBytes, createHmac, timingSafeEqual } from "node:crypto";
import ExcelJS from "exceljs";
import pg from "pg";
import { createCanvas, loadImage, GlobalFonts } from "@napi-rs/canvas";

const {
  SPOTIFY_CLIENT_ID, SPOTIFY_CLIENT_SECRET, DATABASE_URL,
  ADMIN_USUARIO, ADMIN_SENHA, ADMIN_SEGREDO, PORT = 3000,
  SITE_URL = "https://www.minhaconquistacomsenacead.com.br", // site da campanha (para onde o link leva)
  SPOTIFY_PLAYLIST_ID = "", // playlist que recebe as músicas escolhidas (opcional)
} = process.env;
const faltando = ["SPOTIFY_CLIENT_ID", "SPOTIFY_CLIENT_SECRET", "DATABASE_URL", "ADMIN_USUARIO", "ADMIN_SENHA", "ADMIN_SEGREDO"]
  .filter((k) => !process.env[k]);
if (faltando.length) {
  console.error(`Defina as variáveis de ambiente: ${faltando.join(", ")}`);
  process.exit(1);
}

// Conexão com o PostgreSQL. A URL vem do provedor (Neon, Supabase, Render...)
// e normalmente já inclui ?sslmode=require.
const db = new pg.Pool({ connectionString: DATABASE_URL, max: 5 });

const app = express();
app.set("trust proxy", 1); // o Render fica na frente do servidor (IP real e HTTPS)
app.use(cors({
  origin: [
    "https://fabiomrqs.github.io",
    "https://www.minhaconquistacomsenacead.com.br",
    "https://minhaconquistacomsenacead.com.br",
  ],
}));
app.use(express.json({ limit: "20kb" }));
app.use(express.static("public"));

/* ---------------- Spotify ---------------- */

let token = null;
let expiraEm = 0;

async function getToken() {
  if (token && Date.now() < expiraEm) return token;
  const auth = Buffer.from(`${SPOTIFY_CLIENT_ID}:${SPOTIFY_CLIENT_SECRET}`).toString("base64");
  const res = await fetch("https://accounts.spotify.com/api/token", {
    method: "POST",
    headers: { Authorization: `Basic ${auth}`, "Content-Type": "application/x-www-form-urlencoded" },
    body: "grant_type=client_credentials",
  });
  if (!res.ok) throw new Error(`Falha ao obter token: ${res.status}`);
  const data = await res.json();
  token = data.access_token;
  expiraEm = Date.now() + (data.expires_in - 60) * 1000;
  return token;
}

const formatarFaixa = (t) => ({
  id: t.id,
  nome: t.name,
  artistas: t.artists.map((a) => a.name).join(", "),
  album: t.album.name,
  ano: t.album.release_date?.slice(0, 4),
  capa: t.album.images[1]?.url || t.album.images[0]?.url,
  duracaoMs: t.duration_ms,
  explicita: t.explicit,
  preview: t.preview_url,
  link: t.external_urls.spotify,
});

app.get("/api/buscar", async (req, res) => {
  const q = (req.query.q || "").trim();
  if (!q) return res.json([]);
  try {
    const url = `https://api.spotify.com/v1/search?q=${encodeURIComponent(q)}&type=track&limit=10&market=BR`;
    const r = await fetch(url, { headers: { Authorization: `Bearer ${await getToken()}` } });
    if (!r.ok) return res.status(r.status).json({ erro: "O Spotify recusou a busca." });
    const data = await r.json();
    res.json(data.tracks.items.map(formatarFaixa));
  } catch (e) {
    console.error(e);
    res.status(500).json({ erro: "Não foi possível falar com o Spotify agora." });
  }
});

/* ---------------- Participações (mural) ----------------
   Guardadas na tabela "participacoes" do PostgreSQL. */

const EXEMPLOS = [
  ["Ana", "conquista", null, "Tempos Modernos", "Lulu Santos"],
  ["Carlos", "dedicatoria", "Júlia", "Velha Infância", "Tribalistas"],
  ["Bia", "dedicatoria", "Pedro", "Trem-Bala", "Ana Vilela"],
  ["Rafa", "conquista", null, "Vou Festejar", "Beth Carvalho"],
  ["Luísa", "conquista", null, "O Sol", "Vitor Kley"],
  ["João", "dedicatoria", "Marina", "Sá Marina", "Wilson Simonal"],
].map(([nome, tipo, amigo, musica, artistas], i) => ({
  id: `exemplo-${i + 1}`,
  nome, email: null, tipo, amigo,
  musica: { id: null, nome: musica, artistas, capa: null, link: null },
  criadoEm: new Date(Date.UTC(2026, 8, 1 + i)).toISOString(),
}));

// Estados e regiões (para o campo do formulário e os filtros do admin)
const REGIOES = {
  Norte: ["AC", "AP", "AM", "PA", "RO", "RR", "TO"],
  Nordeste: ["AL", "BA", "CE", "MA", "PB", "PE", "PI", "RN", "SE"],
  "Centro-Oeste": ["DF", "GO", "MT", "MS"],
  Sudeste: ["ES", "MG", "RJ", "SP"],
  Sul: ["PR", "RS", "SC"],
};
const UFS = Object.values(REGIOES).flat();
const regiaoDaUf = (uf) => Object.keys(REGIOES).find((r) => REGIOES[r].includes(uf)) || null;

async function inserir(p) {
  await db.query(
    `INSERT INTO participacoes
       (id, nome, email, tipo, amigo, uf, musica_id, musica_nome, musica_artistas, musica_capa, musica_link, criado_em)
     VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12)
     ON CONFLICT (id) DO NOTHING`,
    [p.id, p.nome, p.email, p.tipo, p.amigo, p.uf || null, p.musica.id, p.musica.nome,
     p.musica.artistas, p.musica.capa, p.musica.link, p.criadoEm],
  );
}

// Cria a tabela se ainda não existir. Se estiver vazia, importa o antigo
// data/participacoes.json (quando houver) ou, senão, os 6 exemplos.
async function prepararBanco() {
  await db.query(`
    CREATE TABLE IF NOT EXISTS participacoes (
      id              TEXT PRIMARY KEY,
      nome            TEXT NOT NULL,
      email           TEXT,
      tipo            TEXT NOT NULL CHECK (tipo IN ('conquista', 'dedicatoria')),
      amigo           TEXT,
      musica_id       TEXT,
      musica_nome     TEXT NOT NULL,
      musica_artistas TEXT,
      musica_capa     TEXT,
      musica_link     TEXT,
      criado_em       TIMESTAMPTZ NOT NULL DEFAULT now()
    );
    CREATE INDEX IF NOT EXISTS participacoes_criado_em ON participacoes (criado_em DESC);
    -- coluna adicionada depois: bancos criados antes recebem ela aqui
    ALTER TABLE participacoes ADD COLUMN IF NOT EXISTS uf TEXT;

    -- acessos à página: um contador por dia (horário de Brasília)
    -- configurações guardadas pelo servidor (ex.: autorização do Spotify)
    CREATE TABLE IF NOT EXISTS config (
      chave TEXT PRIMARY KEY,
      valor TEXT NOT NULL
    );

    -- músicas que o servidor já colocou na playlist (evita repetir a mesma faixa)
    CREATE TABLE IF NOT EXISTS playlist_faixas (
      musica_id     TEXT PRIMARY KEY,
      adicionada_em TIMESTAMPTZ NOT NULL DEFAULT now()
    );

    CREATE TABLE IF NOT EXISTS acessos (
      dia   DATE PRIMARY KEY,
      total INTEGER NOT NULL DEFAULT 0
    );
  `);

  const { rows } = await db.query("SELECT count(*)::int AS total FROM participacoes");
  if (rows[0].total > 0) return;

  let iniciais = EXEMPLOS;
  try {
    iniciais = JSON.parse(await readFile("data/participacoes.json", "utf8"));
    console.log(`Importando ${iniciais.length} participações de data/participacoes.json`);
  } catch { /* sem arquivo antigo: usa os exemplos */ }
  for (const p of iniciais) await inserir(p);
}

// Linha do banco → formato que o front-end já usa (sem e-mail)
const publico = (r) => ({
  id: r.id,
  nome: primeiroNome(r.nome || ""),          // o mural só recebe o primeiro nome
  tipo: r.tipo,
  amigo: r.amigo ? primeiroNome(r.amigo) : null,
  musica: { id: r.musica_id, nome: r.musica_nome, artistas: r.musica_artistas, capa: r.musica_capa, link: r.musica_link },
  criadoEm: new Date(r.criado_em).toISOString(),
});
const primeiroNome = (s) => s.trim().split(/\s+/)[0];

app.get("/api/participacoes", async (_req, res) => {
  try {
    // O e-mail nunca entra no SELECT, então nunca sai do servidor
    const { rows } = await db.query(
      `SELECT id, nome, tipo, amigo, musica_id, musica_nome, musica_artistas, musica_capa, musica_link, criado_em
       FROM participacoes ORDER BY criado_em DESC LIMIT 100`,
    );
    res.json(rows.map(publico));
  } catch (e) {
    console.error(e);
    res.status(500).json({ erro: "Não foi possível carregar o mural agora." });
  }
});

app.post("/api/participacoes", async (req, res) => {
  const { nome = "", email = "", tipo, amigo = "", uf = "", aceite, musicaId = "" } = req.body || {};
  const erros = [];
  if (nome.trim().length < 2) erros.push("nome");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim())) erros.push("email");
  if (!["conquista", "dedicatoria"].includes(tipo)) erros.push("tipo");
  if (tipo === "dedicatoria" && amigo.trim().length < 2) erros.push("amigo");
  if (!UFS.includes(uf)) erros.push("uf");
  if (aceite !== true) erros.push("aceite");
  if (!/^[A-Za-z0-9]{22}$/.test(musicaId)) erros.push("musica");
  if (erros.length) return res.status(400).json({ erro: "Dados incompletos.", campos: erros });

  try {
    // Confirma a faixa direto no Spotify em vez de confiar no que veio do navegador
    const r = await fetch(`https://api.spotify.com/v1/tracks/${musicaId}?market=BR`, {
      headers: { Authorization: `Bearer ${await getToken()}` },
    });
    if (!r.ok) return res.status(400).json({ erro: "Música não encontrada no Spotify." });
    const f = formatarFaixa(await r.json());

    const nova = {
      id: randomUUID(),
      nome: primeiroNome(nome).slice(0, 30),
      email: email.trim().toLowerCase(),
      tipo,
      amigo: tipo === "dedicatoria" ? primeiroNome(amigo).slice(0, 30) : null,
      uf,
      musica: { id: f.id, nome: f.nome, artistas: f.artistas, capa: f.capa, link: f.link },
      criadoEm: new Date().toISOString(),
    };
    await inserir(nova);
    // Coloca a música na playlist em segundo plano (a pessoa não espera por isso)
    adicionarNaPlaylist([f.id]).catch((e) => console.error("Playlist:", e.message));
    const { email: _e, uf: _u, ...semEmail } = nova;
    res.status(201).json(semEmail);
  } catch (e) {
    console.error(e);
    res.status(500).json({ erro: "Não foi possível salvar sua participação agora." });
  }
});

/* ---------------- Contador de acessos ----------------
   A página chama esta rota uma vez por visita. Não guarda IP nem nada pessoal,
   só soma 1 no contador do dia. */

const ultimoAcesso = new Map(); // freio simples contra recarregar sem parar
app.post("/api/acesso", async (req, res) => {
  const agora = Date.now();
  if (agora - (ultimoAcesso.get(req.ip) || 0) < 10_000) return res.status(204).end();
  ultimoAcesso.set(req.ip, agora);
  if (ultimoAcesso.size > 5000) ultimoAcesso.clear();
  try {
    await db.query(
      `INSERT INTO acessos (dia, total) VALUES ((now() AT TIME ZONE 'America/Sao_Paulo')::date, 1)
       ON CONFLICT (dia) DO UPDATE SET total = acessos.total + 1`,
    );
  } catch (e) {
    console.error(e);
  }
  res.status(204).end();
});

/* ---------------- Compartilhamento (Facebook e WhatsApp) ----------------
   Facebook e WhatsApp não rodam o JavaScript da página: eles leem as tags og:
   do HTML. Por isso cada participação ganha um link próprio (/c/<id>) com as
   tags apontando para uma imagem gerada aqui, com a capa da música.
   Quem clica no link é levado para o site da campanha. */

const PASTA_FONTES = new URL("./fonts/", import.meta.url).pathname;
GlobalFonts.registerFromPath(`${PASTA_FONTES}Montserrat-ExtraBold.ttf`, "Montserrat ExtraBold");
GlobalFonts.registerFromPath(`${PASTA_FONTES}Montserrat-Bold.ttf`, "Montserrat Bold");
GlobalFonts.registerFromPath(`${PASTA_FONTES}Montserrat-Regular.ttf`, "Montserrat");

const escHtml = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
const urlPublica = (req) => `${req.protocol}://${req.get("host")}`;

async function buscarParticipacao(id) {
  if (!/^[A-Za-z0-9-]{1,64}$/.test(id)) return null;
  const { rows } = await db.query(
    `SELECT id, nome, tipo, amigo, musica_nome, musica_artistas, musica_capa FROM participacoes WHERE id = $1`, [id],
  );
  return rows[0] || null;
}

const tituloDe = (p) => (p.tipo === "dedicatoria" ? `A música da conquista de ${primeiroNome(p.amigo || "")} é` : "A música da minha conquista é");

async function carregarCapa(url) {
  if (!url) return null;
  const grande = url.replace("ab67616d00001e02", "ab67616d0000b273"); // versão 640px do Spotify
  for (const u of [grande, url]) {
    try {
      const r = await fetch(u);
      if (r.ok) return await loadImage(Buffer.from(await r.arrayBuffer()));
    } catch { /* tenta a próxima */ }
  }
  return null;
}

function quebrar(ctx, texto, largura, maxLinhas) {
  const linhas = [];
  let linha = "";
  for (const p of String(texto).split(/\s+/).filter(Boolean)) {
    const t = linha ? `${linha} ${p}` : p;
    if (ctx.measureText(t).width <= largura || !linha) linha = t;
    else { linhas.push(linha); linha = p; }
  }
  if (linha) linhas.push(linha);
  if (linhas.length > maxLinhas) {
    linhas.length = maxLinhas;
    let u = linhas[maxLinhas - 1];
    while (u && ctx.measureText(`${u}…`).width > largura) u = u.slice(0, -1);
    linhas[maxLinhas - 1] = `${u.trimEnd()}…`;
  }
  return linhas;
}

function arredondado(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.roundRect(x, y, w, h, r);
}

// Nota musical simples (quando a música não tem capa)
function desenharNota(ctx, cx, cy, t) {
  ctx.save();
  ctx.fillStyle = "#fff"; ctx.strokeStyle = "#fff";
  ctx.lineWidth = t * .09; ctx.lineCap = "round";
  const xe = cx - t * .3, xd = cx + t * .3, yb = cy + t * .32, yt = cy - t * .42;
  ctx.beginPath(); ctx.moveTo(xe + t * .14, yb); ctx.lineTo(xe + t * .14, yt + t * .12);
  ctx.lineTo(xd + t * .14, yt); ctx.lineTo(xd + t * .14, yb - t * .12); ctx.stroke();
  for (const [x, y] of [[xe, yb], [xd, yb - t * .12]]) {
    ctx.beginPath(); ctx.ellipse(x, y, t * .17, t * .13, -.4, 0, Math.PI * 2); ctx.fill();
  }
  ctx.restore();
}

// Imagem 1200×630, o tamanho recomendado pelo Facebook e aceito pelo WhatsApp
async function gerarImagemOg({ titulo, musica, artistas, capaUrl }) {
  const L = 1200, A = 630;
  const canvas = createCanvas(L, A);
  const ctx = canvas.getContext("2d");
  const capa = await carregarCapa(capaUrl);

  // Fundo branco com a mancha azul do site
  ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, L, A);
  const g = ctx.createRadialGradient(L * .3, A * .5, 0, L * .3, A * .5, L * .6);
  g.addColorStop(0, "rgba(71,151,203,.18)"); g.addColorStop(1, "rgba(71,151,203,0)");
  ctx.fillStyle = g; ctx.fillRect(0, 0, L, A);

  // Cartão azul
  const cx = 36, cy = 36, cw = L - 72, ch = A - 72;
  ctx.save();
  ctx.shadowColor = "rgba(29,68,128,.25)"; ctx.shadowBlur = 30; ctx.shadowOffsetY = 10;
  arredondado(ctx, cx, cy, cw, ch, 36); ctx.fillStyle = "#4797CB"; ctx.fill();
  ctx.restore();

  // Capa à esquerda
  const tc = 440, xc = cx + 52, yc = (A - tc) / 2;
  ctx.save();
  arredondado(ctx, xc, yc, tc, tc, 24); ctx.clip();
  ctx.fillStyle = "#1D4480"; ctx.fillRect(xc, yc, tc, tc);
  if (capa) ctx.drawImage(capa, xc, yc, tc, tc);
  else desenharNota(ctx, xc + tc / 2, yc + tc / 2, tc * .42); // sem capa: nota musical
  ctx.restore();

  // Textos à direita
  const xt = xc + tc + 52, lt = cx + cw - 52 - xt;
  ctx.textBaseline = "top"; ctx.fillStyle = "#fff";
  ctx.font = "30px 'Montserrat ExtraBold'";
  const lTit = quebrar(ctx, titulo.toUpperCase(), lt, 2);
  let tam = 58, lMus;
  do { ctx.font = `${tam}px 'Montserrat Bold'`; lMus = quebrar(ctx, musica.toUpperCase(), lt, 3); tam -= 4; }
  while (lMus.some((l) => l.endsWith("…")) && tam >= 42);
  const tamMus = tam + 4;
  ctx.font = "30px Montserrat";
  const lArt = artistas ? quebrar(ctx, artistas, lt, 2) : [];

  const altura = lTit.length * 38 + 28 + lMus.length * (tamMus * 1.1) + 14 + lArt.length * 38;
  let y = cy + (ch - altura) / 2;

  ctx.font = "30px 'Montserrat ExtraBold'";
  lTit.forEach((t) => { ctx.fillText(t, xt, y); y += 38; });
  y += 28;
  ctx.font = `${tamMus}px 'Montserrat Bold'`;
  lMus.forEach((t) => { ctx.fillText(t, xt, y); y += tamMus * 1.1; });
  y += 14;
  ctx.globalAlpha = .92; ctx.font = "30px Montserrat";
  lArt.forEach((t) => { ctx.fillText(t, xt, y); y += 38; });
  ctx.globalAlpha = 1;

  return canvas.encode("jpeg", 88); // JPEG fica leve (o WhatsApp ignora imagens pesadas)
}

// Guarda as últimas imagens geradas para não refazer a cada visita do Facebook/WhatsApp
const cacheImagens = new Map();
async function imagemEmCache(chave, gerar) {
  if (cacheImagens.has(chave)) return cacheImagens.get(chave);
  const buf = await gerar();
  cacheImagens.set(chave, buf);
  if (cacheImagens.size > 300) cacheImagens.delete(cacheImagens.keys().next().value);
  return buf;
}

function enviarJpeg(res, buf) {
  res.set("Content-Type", "image/jpeg");
  res.set("Cache-Control", "public, max-age=604800"); // 7 dias
  res.send(buf);
}

app.get("/c/:id/imagem.jpg", async (req, res) => {
  try {
    const p = await buscarParticipacao(req.params.id);
    if (!p) return res.status(404).end();
    const buf = await imagemEmCache(p.id, () => gerarImagemOg({
      titulo: tituloDe(p), musica: p.musica_nome, artistas: p.musica_artistas, capaUrl: p.musica_capa,
    }));
    enviarJpeg(res, buf);
  } catch (e) {
    console.error(e);
    res.status(500).end();
  }
});

// Imagem genérica da campanha (para a página inicial do site)
app.get("/og/campanha.jpg", async (_req, res) => {
  try {
    const buf = await imagemEmCache("campanha", () => gerarImagemOg({
      titulo: "Senac EAD", musica: "Qual é a música da sua conquista?", artistas: "Escolha a sua e compartilhe!",
    }));
    enviarJpeg(res, buf);
  } catch (e) {
    console.error(e);
    res.status(500).end();
  }
});

app.get("/c/:id", async (req, res) => {
  try {
    const p = await buscarParticipacao(req.params.id);
    if (!p) return res.redirect(302, SITE_URL);
    const base = urlPublica(req);
    const url = `${base}/c/${encodeURIComponent(p.id)}`;
    const titulo = `${tituloDe(p)} “${p.musica_nome}”`;
    const descricao = `${p.musica_artistas ? `${p.musica_artistas}. ` : ""}Escolha a música da sua conquista também!`;
    res.set("Cache-Control", "public, max-age=300");
    res.type("html").send(`<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escHtml(titulo)} | Senac EAD</title>
<meta name="description" content="${escHtml(descricao)}">
<meta property="og:type" content="website">
<meta property="og:locale" content="pt_BR">
<meta property="og:site_name" content="Qual é a música da sua conquista? | Senac EAD">
<meta property="og:url" content="${escHtml(url)}">
<meta property="og:title" content="${escHtml(titulo)}">
<meta property="og:description" content="${escHtml(descricao)}">
<meta property="og:image" content="${escHtml(`${url}/imagem.jpg`)}">
<meta property="og:image:secure_url" content="${escHtml(`${url}/imagem.jpg`)}">
<meta property="og:image:type" content="image/jpeg">
<meta property="og:image:width" content="1200">
<meta property="og:image:height" content="630">
<meta property="og:image:alt" content="${escHtml(`Capa de ${p.musica_nome}`)}">
<meta name="twitter:card" content="summary_large_image">
<script>location.replace(${JSON.stringify(SITE_URL)});</script>
</head>
<body style="font-family:system-ui,sans-serif;text-align:center;padding:3rem 1rem">
<p><a href="${escHtml(SITE_URL)}">Escolha a música da sua conquista no site do Senac EAD</a></p>
</body>
</html>`);
  } catch (e) {
    console.error(e);
    res.redirect(302, SITE_URL);
  }
});

/* ---------------- Admin ----------------
   Login simples com um usuário e senha definidos nas variáveis de ambiente.
   Depois do login, o navegador recebe um cookie assinado válido por 8 horas. */

const COOKIE = "admin_sessao";
const DURACAO_SESSAO = 8 * 60 * 60 * 1000;

const assinar = (texto) => createHmac("sha256", ADMIN_SEGREDO).update(texto).digest("base64url");
const iguais = (a, b) => {
  const x = Buffer.from(String(a)), y = Buffer.from(String(b));
  return x.length === y.length && timingSafeEqual(x, y);
};

function criarSessao() {
  const expira = String(Date.now() + DURACAO_SESSAO);
  return `${expira}.${assinar(expira)}`;
}

function sessaoValida(req) {
  const cookie = (req.headers.cookie || "").split(/;\s*/).find((c) => c.startsWith(`${COOKIE}=`));
  if (!cookie) return false;
  const [expira, assinatura] = cookie.slice(COOKIE.length + 1).split(".");
  return Boolean(expira && assinatura) && iguais(assinatura, assinar(expira)) && Date.now() < Number(expira);
}

function exigirLogin(req, res, next) {
  if (sessaoValida(req)) return next();
  res.status(401).json({ erro: "Faça login novamente." });
}

// Limita tentativas de senha: 5 erros por IP a cada 15 minutos
const tentativas = new Map();
function bloqueado(ip) {
  const t = tentativas.get(ip);
  if (!t || Date.now() > t.ate) return false;
  return t.erros >= 5;
}
function registrarErro(ip) {
  const t = tentativas.get(ip);
  if (!t || Date.now() > t.ate) tentativas.set(ip, { erros: 1, ate: Date.now() + 15 * 60 * 1000 });
  else t.erros++;
}

app.get("/admin", (_req, res) => {
  res.set("Cache-Control", "no-store");
  res.sendFile(new URL("./admin.html", import.meta.url).pathname);
});

app.post("/api/admin/login", (req, res) => {
  const ip = req.ip;
  if (bloqueado(ip)) return res.status(429).json({ erro: "Muitas tentativas. Espere 15 minutos." });
  const { usuario = "", senha = "" } = req.body || {};
  if (!iguais(usuario, ADMIN_USUARIO) || !iguais(senha, ADMIN_SENHA)) {
    registrarErro(ip);
    return res.status(401).json({ erro: "Usuário ou senha incorretos." });
  }
  tentativas.delete(ip);
  res.cookie(COOKIE, criarSessao(), {
    httpOnly: true, sameSite: "strict", secure: req.secure, maxAge: DURACAO_SESSAO, path: "/",
  });
  res.json({ ok: true });
});

app.post("/api/admin/logout", (_req, res) => {
  res.clearCookie(COOKIE, { path: "/" });
  res.json({ ok: true });
});

app.get("/api/admin/sessao", (req, res) => res.json({ logado: sessaoValida(req) }));

// Monta o WHERE a partir dos filtros da tela (estados, tipo e período).
// Os 6 exemplos do mural nunca entram no admin.
function filtros(q) {
  const cond = ["id NOT LIKE 'exemplo-%'"], valores = [];
  const add = (sql, v) => { valores.push(v); cond.push(sql.replace("?", `$${valores.length}`)); };

  // uf=SP,RJ,nao-informado (um ou mais estados)
  const escolhidos = String(q.uf || "").split(",").filter(Boolean);
  const ufs = escolhidos.filter((u) => UFS.includes(u));
  const semUf = escolhidos.includes("nao-informado");
  if (ufs.length || semUf) {
    const partes = [];
    if (ufs.length) { valores.push(ufs); partes.push(`uf = ANY($${valores.length})`); }
    if (semUf) partes.push("uf IS NULL");
    cond.push(`(${partes.join(" OR ")})`);
  }

  if (["conquista", "dedicatoria"].includes(q.tipo)) add("tipo = ?", q.tipo);
  if (/^\d{4}-\d{2}-\d{2}$/.test(q.de || "")) add("criado_em >= (?::date::timestamp AT TIME ZONE 'America/Sao_Paulo')", q.de);
  if (/^\d{4}-\d{2}-\d{2}$/.test(q.ate || "")) add("criado_em < ((?::date + 1)::timestamp AT TIME ZONE 'America/Sao_Paulo')", q.ate);

  return { where: `WHERE ${cond.join(" AND ")}`, valores };
}

/* ---------------- Playlist automática no Spotify ----------------
   Para mexer numa playlist, o Spotify exige a autorização da conta dona dela
   (a chave do app sozinha só permite buscar músicas). Essa autorização é feita
   uma única vez em /admin/spotify e fica guardada na tabela "config". */

const ESCOPOS = "playlist-modify-public playlist-modify-private";
const COOKIE_ESTADO = "spotify_estado";
let tokenUsuario = null, tokenUsuarioExpira = 0;

const lerConfig = async (chave) => (await db.query("SELECT valor FROM config WHERE chave = $1", [chave])).rows[0]?.valor || null;
const salvarConfig = (chave, valor) => db.query(
  "INSERT INTO config (chave, valor) VALUES ($1, $2) ON CONFLICT (chave) DO UPDATE SET valor = EXCLUDED.valor", [chave, valor],
);

async function pedirTokenSpotify(corpo) {
  const auth = Buffer.from(`${SPOTIFY_CLIENT_ID}:${SPOTIFY_CLIENT_SECRET}`).toString("base64");
  const r = await fetch("https://accounts.spotify.com/api/token", {
    method: "POST",
    headers: { Authorization: `Basic ${auth}`, "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(corpo),
  });
  const dados = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(`Spotify recusou o token (${r.status} ${dados.error || ""})`.trim());
  return dados;
}

// Token da conta dona da playlist (renovado sozinho quando expira)
async function getTokenUsuario() {
  if (tokenUsuario && Date.now() < tokenUsuarioExpira) return tokenUsuario;
  const refresh = await lerConfig("spotify_refresh_token");
  if (!refresh) return null;
  const d = await pedirTokenSpotify({ grant_type: "refresh_token", refresh_token: refresh });
  if (d.refresh_token) await salvarConfig("spotify_refresh_token", d.refresh_token);
  tokenUsuario = d.access_token;
  tokenUsuarioExpira = Date.now() + (d.expires_in - 60) * 1000;
  return tokenUsuario;
}

async function spotifyPlaylist(metodo, corpo) {
  const token = await getTokenUsuario();
  if (!token) throw new Error("conta do Spotify ainda não conectada em /admin/spotify");
  const r = await fetch(`https://api.spotify.com/v1/playlists/${encodeURIComponent(SPOTIFY_PLAYLIST_ID)}/tracks`, {
    method: metodo,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(corpo),
  });
  if (!r.ok) {
    const d = await r.json().catch(() => ({}));
    throw new Error(`Spotify respondeu ${r.status}: ${d.error?.message || "erro desconhecido"}`);
  }
}

// Adiciona as faixas que ainda não estão na playlist (no máximo 100 por pedido, regra do Spotify)
async function adicionarNaPlaylist(ids) {
  if (!SPOTIFY_PLAYLIST_ID) return 0;
  const unicos = [...new Set(ids.filter((id) => /^[A-Za-z0-9]{22}$/.test(id)))];
  let adicionadas = 0;
  for (let i = 0; i < unicos.length; i += 100) {
    // "Reserva" as faixas antes de enviar: se duas pessoas escolherem a mesma música ao mesmo tempo, só uma envia
    const { rows } = await db.query(
      `INSERT INTO playlist_faixas (musica_id) SELECT unnest($1::text[]) ON CONFLICT DO NOTHING RETURNING musica_id`,
      [unicos.slice(i, i + 100)],
    );
    const novas = rows.map((r) => r.musica_id);
    if (!novas.length) continue;
    try {
      await spotifyPlaylist("POST", { uris: novas.map((id) => `spotify:track:${id}`) });
      adicionadas += novas.length;
    } catch (e) {
      await db.query("DELETE FROM playlist_faixas WHERE musica_id = ANY($1)", [novas]); // tenta de novo na próxima
      throw e;
    }
  }
  return adicionadas;
}

async function removerDaPlaylistSeSobrou(musicaId) {
  if (!SPOTIFY_PLAYLIST_ID) return;
  const { rows } = await db.query("SELECT 1 FROM participacoes WHERE musica_id = $1 LIMIT 1", [musicaId]);
  if (rows.length) return; // outra pessoa também escolheu essa música
  const { rowCount } = await db.query("DELETE FROM playlist_faixas WHERE musica_id = $1", [musicaId]);
  if (rowCount) await spotifyPlaylist("DELETE", { tracks: [{ uri: `spotify:track:${musicaId}` }] });
}

// Página simples de configuração (só para quem está logado no admin)
function paginaSpotify(res, { conta, naPlaylist, escolhidas, mensagem, erro }) {
  const pronta = Boolean(SPOTIFY_PLAYLIST_ID);
  const faltam = Math.max(0, escolhidas - naPlaylist);
  res.set("Cache-Control", "no-store");
  res.type("html").send(`<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex"><title>Playlist do Spotify | Painel</title>
<style>
  body { margin: 0; background: #F3F6FA; color: #2B3A55; font: 16px/1.5 "Montserrat", system-ui, sans-serif; }
  main { max-width: 620px; margin: 2.5rem auto; padding: 0 1.25rem; }
  .card { background: #fff; border-radius: 18px; padding: 1.75rem; }
  h1 { color: #1D4480; font-size: 1.4rem; margin: 0 0 .25rem; }
  .voltar { color: #4797CB; font-weight: 700; text-decoration: none; display: inline-block; margin-bottom: 1rem; }
  dl { display: grid; grid-template-columns: auto 1fr; gap: .5rem 1rem; margin: 1.25rem 0; }
  dt { color: #5E6B7A; font-weight: 600; } dd { margin: 0; font-weight: 700; word-break: break-all; }
  .ok { color: #2E7D4F; } .nao { color: #C0392B; }
  .msg { padding: .8rem 1rem; border-radius: 10px; background: #EAF3FA; margin: 1rem 0; }
  .msg.erro { background: #FCEDEB; color: #A93226; }
  .acoes { display: flex; gap: .6rem; flex-wrap: wrap; margin-top: 1.25rem; }
  .btn { display: inline-flex; align-items: center; min-height: 44px; padding: .6rem 1.2rem; border-radius: 999px; border: 2px solid #1D4480;
    background: #1D4480; color: #fff; font: inherit; font-weight: 700; line-height: 1.2; text-decoration: none; cursor: pointer; }
  dd a { color: #1D4480; }
  .btn.contorno { background: #fff; color: #1D4480; }
  p.dica { color: #5E6B7A; font-size: .9rem; }
</style></head><body><main>
<a class="voltar" href="/admin">← Voltar ao painel</a>
<div class="card">
  <h1>Playlist do Spotify</h1>
  <p class="dica">Cada música escolhida no site entra sozinha na playlist. Se uma participação for removida no painel e ninguém mais tiver escolhido aquela música, ela sai da playlist também.</p>
  ${mensagem ? `<p class="msg">${escHtml(mensagem)}</p>` : ""}
  ${erro ? `<p class="msg erro">${escHtml(erro)}</p>` : ""}
  <dl>
    <dt>Playlist</dt><dd>${pronta ? `<a href="https://open.spotify.com/playlist/${escHtml(SPOTIFY_PLAYLIST_ID)}" target="_blank" rel="noopener">${escHtml(SPOTIFY_PLAYLIST_ID)}</a>` : `<span class="nao">defina SPOTIFY_PLAYLIST_ID no Render</span>`}</dd>
    <dt>Conta conectada</dt><dd>${conta ? `<span class="ok">${escHtml(conta)}</span>` : `<span class="nao">nenhuma</span>`}</dd>
    <dt>Músicas escolhidas</dt><dd>${escolhidas}</dd>
    <dt>Já na playlist</dt><dd>${naPlaylist}${faltam ? ` <span class="nao">(faltam ${faltam})</span>` : ""}</dd>
  </dl>
  <div class="acoes">
    ${pronta ? `<a class="btn ${conta ? "contorno" : ""}" href="/admin/spotify/conectar">${conta ? "Conectar outra conta" : "Conectar conta do Spotify"}</a>` : ""}
    ${pronta && conta && faltam ? `<form method="post" action="/admin/spotify/sincronizar"><button class="btn">Adicionar as ${faltam} que faltam</button></form>` : ""}
  </div>
  ${pronta ? `<p class="dica">Conecte com a conta <b>dona da playlist</b> (ou que seja colaboradora dela).</p>` : ""}
</div></main></body></html>`);
}

async function statusSpotify() {
  const [conta, na, esc] = await Promise.all([
    lerConfig("spotify_conta"),
    db.query("SELECT count(*)::int AS n FROM playlist_faixas"),
    db.query("SELECT count(DISTINCT musica_id)::int AS n FROM participacoes WHERE musica_id IS NOT NULL"),
  ]);
  return { conta, naPlaylist: na.rows[0].n, escolhidas: esc.rows[0].n };
}

const exigirLoginPagina = (req, res, next) => (sessaoValida(req) ? next() : res.redirect("/admin"));
const urlRetornoSpotify = (req) => `${urlPublica(req)}/admin/spotify/callback`;

app.get("/admin/spotify", exigirLoginPagina, async (req, res) => {
  const mensagem = req.query.status === "ok" ? "Conta conectada! As próximas músicas já vão para a playlist." : req.query.msg;
  paginaSpotify(res, { ...(await statusSpotify()), mensagem, erro: req.query.erro });
});

app.get("/admin/spotify/conectar", exigirLoginPagina, (req, res) => {
  const estado = randomBytes(16).toString("hex");
  res.cookie(COOKIE_ESTADO, estado, { httpOnly: true, secure: req.secure, sameSite: "lax", maxAge: 10 * 60 * 1000, path: "/admin/spotify" });
  const p = new URLSearchParams({
    client_id: SPOTIFY_CLIENT_ID, response_type: "code", redirect_uri: urlRetornoSpotify(req),
    scope: ESCOPOS, state: estado, show_dialog: "true",
  });
  res.redirect(`https://accounts.spotify.com/authorize?${p}`);
});

// Sem exigir o login aqui: o cookie do admin é "Strict" e não vem quando o Spotify
// redireciona de volta. A proteção é o código de estado, criado só por quem estava logado.
app.get("/admin/spotify/callback", async (req, res) => {
  const cookie = (req.headers.cookie || "").split(/;\s*/).find((c) => c.startsWith(`${COOKIE_ESTADO}=`))?.split("=")[1];
  res.clearCookie(COOKIE_ESTADO, { path: "/admin/spotify" });
  const voltar = (q) => res.redirect(`/admin/spotify?${new URLSearchParams(q)}`);
  if (req.query.error) return voltar({ erro: "A autorização foi cancelada no Spotify." });
  if (!cookie || !req.query.state || !iguais(cookie, req.query.state)) return voltar({ erro: "A autorização expirou. Tente conectar de novo." });
  try {
    const d = await pedirTokenSpotify({ grant_type: "authorization_code", code: String(req.query.code || ""), redirect_uri: urlRetornoSpotify(req) });
    if (!d.refresh_token) throw new Error("o Spotify não enviou a autorização permanente");
    const eu = await fetch("https://api.spotify.com/v1/me", { headers: { Authorization: `Bearer ${d.access_token}` } }).then((r) => r.json());
    await salvarConfig("spotify_refresh_token", d.refresh_token);
    await salvarConfig("spotify_conta", eu.display_name || eu.id || "conta do Spotify");
    tokenUsuario = d.access_token;
    tokenUsuarioExpira = Date.now() + (d.expires_in - 60) * 1000;
    voltar({ status: "ok" });
  } catch (e) {
    console.error(e);
    voltar({ erro: `Não foi possível conectar: ${e.message}` });
  }
});

// Coloca na playlist as músicas escolhidas antes da conexão (ou que falharam)
app.post("/admin/spotify/sincronizar", exigirLoginPagina, async (req, res) => {
  try {
    const { rows } = await db.query(
      `SELECT musica_id FROM participacoes WHERE musica_id IS NOT NULL
       GROUP BY musica_id ORDER BY min(criado_em)`,
    );
    const n = await adicionarNaPlaylist(rows.map((r) => r.musica_id));
    res.redirect(`/admin/spotify?${new URLSearchParams({ msg: `${n} ${n === 1 ? "música adicionada" : "músicas adicionadas"} à playlist.` })}`);
  } catch (e) {
    console.error(e);
    res.redirect(`/admin/spotify?${new URLSearchParams({ erro: `Não foi possível adicionar: ${e.message}` })}`);
  }
});

const POR_PAGINA = 50;

// Ordenação da tabela: só colunas desta lista (nunca texto vindo direto do navegador)
const ORDENS = {
  data: "criado_em",
  nome: "lower(nome)",
  email: "lower(email)",
  tipo: "tipo",
  uf: "uf",
  musica: "lower(musica_nome)",
};
function ordenacao(q) {
  const coluna = ORDENS[q.ordem] || ORDENS.data;
  const direcao = q.direcao === "asc" ? "ASC" : "DESC";
  return `ORDER BY ${coluna} ${direcao} NULLS LAST, criado_em DESC`;
}

app.get("/api/admin/painel", exigirLogin, async (req, res) => {
  const { where, valores } = filtros(req.query);
  const pagina = Math.max(1, parseInt(req.query.pagina, 10) || 1);
  try {
    const [geral, filtrado, linhas, musicas, artistas, acessos] = await Promise.all([
      db.query(`SELECT count(*)::int AS total FROM participacoes WHERE id NOT LIKE 'exemplo-%'`),
      db.query(`SELECT count(*)::int AS total FROM participacoes ${where}`, valores),
      db.query(`SELECT id, nome, email, tipo, amigo, uf, musica_nome, musica_artistas, criado_em
                FROM participacoes ${where} ${ordenacao(req.query)}
                LIMIT ${POR_PAGINA} OFFSET ${(pagina - 1) * POR_PAGINA}`, valores),
      db.query(`SELECT musica_nome AS nome, musica_artistas AS artistas, count(*)::int AS total
                FROM participacoes ${where} GROUP BY 1, 2 ORDER BY total DESC, nome LIMIT 10`, valores),
      db.query(`SELECT musica_artistas AS nome, count(*)::int AS total
                FROM participacoes ${where} GROUP BY 1 ORDER BY total DESC, nome LIMIT 10`, valores),
      db.query(`SELECT coalesce(sum(total), 0)::int AS total,
                       coalesce(sum(total) FILTER (WHERE dia = (now() AT TIME ZONE 'America/Sao_Paulo')::date), 0)::int AS hoje
                FROM acessos`),
    ]);
    const total = filtrado.rows[0].total;
    res.json({
      atualizadoEm: new Date().toISOString(),
      totalGeral: geral.rows[0].total,
      acessosTotal: acessos.rows[0].total,
      acessosHoje: acessos.rows[0].hoje,
      total,
      pagina,
      paginas: Math.max(1, Math.ceil(total / POR_PAGINA)),
      porPagina: POR_PAGINA,
      participantes: linhas.rows,
      musicas: musicas.rows,
      artistas: artistas.rows,
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ erro: "Não foi possível carregar o painel." });
  }
});

// Remove uma participação de vez: some do admin, do mural e o link de compartilhar deixa de funcionar
app.delete("/api/admin/participacoes/:id", exigirLogin, async (req, res) => {
  try {
    const { rows } = await db.query("DELETE FROM participacoes WHERE id = $1 RETURNING musica_id", [req.params.id]);
    cacheImagens.delete(req.params.id);
    if (!rows.length) return res.status(404).json({ erro: "Essa participação já tinha sido removida." });
    // Se ninguém mais escolheu essa música, ela sai da playlist também
    if (rows[0].musica_id) removerDaPlaylistSeSobrou(rows[0].musica_id).catch((e) => console.error("Playlist:", e.message));
    res.status(204).end();
  } catch (e) {
    console.error(e);
    res.status(500).json({ erro: "Não foi possível remover agora." });
  }
});

app.get("/api/admin/exportar", exigirLogin, async (req, res) => {
  const formato = req.query.formato === "xlsx" ? "xlsx" : "csv";
  const { where, valores } = filtros(req.query);
  try {
    const { rows } = await db.query(
      `SELECT * FROM participacoes ${where} ${ordenacao(req.query)}`, valores,
    );
    const colunas = [
      ["Data", (r) => new Date(r.criado_em).toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" })],
      ["Nome", (r) => r.nome],
      ["E-mail", (r) => r.email || ""],
      ["Tipo", (r) => (r.tipo === "dedicatoria" ? "Dedicatória" : "Conquista")],
      ["Dedicada a", (r) => r.amigo || ""],
      ["Estado", (r) => r.uf || "Não informado"],
      ["Região", (r) => regiaoDaUf(r.uf) || "Não informado"],
      ["Música", (r) => r.musica_nome],
      ["Artistas", (r) => r.musica_artistas || ""],
      ["Link no Spotify", (r) => r.musica_link || ""],
    ];
    const nomeArquivo = `participantes-${new Date().toISOString().slice(0, 10)}.${formato}`;
    res.set("Content-Disposition", `attachment; filename="${nomeArquivo}"`);
    res.set("Cache-Control", "no-store");

    if (formato === "csv") {
      // ";" e BOM para o Excel em português abrir com acentos e colunas certas
      const celula = (v) => {
        let t = String(v ?? "");
        if (/^[=+\-@]/.test(t)) t = `'${t}`; // evita fórmulas maliciosas ao abrir no Excel
        return /[";\n]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t;
      };
      const linhas = [colunas.map(([t]) => t), ...rows.map((r) => colunas.map(([, f]) => f(r)))];
      res.type("text/csv; charset=utf-8");
      return res.send("\uFEFF" + linhas.map((l) => l.map(celula).join(";")).join("\r\n"));
    }

    const livro = new ExcelJS.Workbook();
    const aba = livro.addWorksheet("Participantes");
    aba.columns = colunas.map(([header]) => ({ header, width: header === "E-mail" || header === "Música" ? 32 : 18 }));
    rows.forEach((r) => aba.addRow(colunas.map(([, f]) => f(r))));
    aba.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    aba.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1D4480" } };
    aba.views = [{ state: "frozen", ySplit: 1 }];
    aba.autoFilter = { from: "A1", to: { row: 1, column: colunas.length } };
    res.type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.send(Buffer.from(await livro.xlsx.writeBuffer()));
  } catch (e) {
    console.error(e);
    res.status(500).json({ erro: "Não foi possível gerar a planilha." });
  }
});

try {
  await prepararBanco();
} catch (e) {
  console.error("Não foi possível conectar ao banco de dados:", e.message);
  process.exit(1);
}

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Servidor rodando na porta ${PORT}`);
});