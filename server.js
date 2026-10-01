import cors from "cors";
import express from "express";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { randomUUID } from "node:crypto";

const { SPOTIFY_CLIENT_ID, SPOTIFY_CLIENT_SECRET, PORT = 3000 } = process.env;
if (!SPOTIFY_CLIENT_ID || !SPOTIFY_CLIENT_SECRET) {
  console.error("Defina SPOTIFY_CLIENT_ID e SPOTIFY_CLIENT_SECRET no arquivo .env");
  process.exit(1);
}

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
   Guardadas em data/participacoes.json. Para produção, troque
   lerTodas/salvarTodas por um banco de dados. */

const ARQUIVO = "data/participacoes.json";

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

async function lerTodas() {
  try {
    return JSON.parse(await readFile(ARQUIVO, "utf8"));
  } catch {
    await salvarTodas(EXEMPLOS);
    return [...EXEMPLOS];
  }
}

async function salvarTodas(lista) {
  await mkdir("data", { recursive: true });
  await writeFile(ARQUIVO, JSON.stringify(lista, null, 2));
}

// Só dados públicos vão para o mural (o e-mail nunca sai do servidor)
const publico = ({ email, ...p }) => p;
const primeiroNome = (s) => s.trim().split(/\s+/)[0];

app.get("/api/participacoes", async (_req, res) => {
  const lista = await lerTodas();
  res.json(lista.sort((a, b) => b.criadoEm.localeCompare(a.criadoEm)).slice(0, 100).map(publico));
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
    const lista = await lerTodas();
    lista.push(nova);
    await salvarTodas(lista);
    res.status(201).json(publico(nova));
  } catch (e) {
    console.error(e);
    res.status(500).json({ erro: "Não foi possível salvar sua participação agora." });
  }
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Servidor rodando na porta ${PORT}`);
});