import cors from "cors";
import express from "express";
import { readFile } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import pg from "pg";

const { SPOTIFY_CLIENT_ID, SPOTIFY_CLIENT_SECRET, DATABASE_URL, PORT = 3000 } = process.env;
if (!SPOTIFY_CLIENT_ID || !SPOTIFY_CLIENT_SECRET || !DATABASE_URL) {
  console.error("Defina SPOTIFY_CLIENT_ID, SPOTIFY_CLIENT_SECRET e DATABASE_URL no arquivo .env");
  process.exit(1);
}

// Conexão com o PostgreSQL. A URL vem do provedor (Neon, Supabase, Render...)
// e normalmente já inclui ?sslmode=require.
const db = new pg.Pool({ connectionString: DATABASE_URL, max: 5 });

const app = express();
app.use(cors({
  origin: "https://fabiomrqs.github.io"
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

async function inserir(p) {
  await db.query(
    `INSERT INTO participacoes
       (id, nome, email, tipo, amigo, musica_id, musica_nome, musica_artistas, musica_capa, musica_link, criado_em)
     VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)
     ON CONFLICT (id) DO NOTHING`,
    [p.id, p.nome, p.email, p.tipo, p.amigo, p.musica.id, p.musica.nome,
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
  nome: r.nome,
  tipo: r.tipo,
  amigo: r.amigo,
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
  const { nome = "", email = "", tipo, amigo = "", aceite, musicaId = "" } = req.body || {};
  const erros = [];
  if (nome.trim().length < 2) erros.push("nome");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim())) erros.push("email");
  if (!["conquista", "dedicatoria"].includes(tipo)) erros.push("tipo");
  if (tipo === "dedicatoria" && amigo.trim().length < 2) erros.push("amigo");
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
      musica: { id: f.id, nome: f.nome, artistas: f.artistas, capa: f.capa, link: f.link },
      criadoEm: new Date().toISOString(),
    };
    await inserir(nova);
    const { email: _, ...semEmail } = nova;
    res.status(201).json(semEmail);
  } catch (e) {
    console.error(e);
    res.status(500).json({ erro: "Não foi possível salvar sua participação agora." });
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