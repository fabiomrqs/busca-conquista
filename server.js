import cors from "cors";
import express from "express";
import { readFile } from "node:fs/promises";
import { randomUUID, createHmac, timingSafeEqual } from "node:crypto";
import ExcelJS from "exceljs";
import pg from "pg";

const {
  SPOTIFY_CLIENT_ID, SPOTIFY_CLIENT_SECRET, DATABASE_URL,
  ADMIN_USUARIO, ADMIN_SENHA, ADMIN_SEGREDO, PORT = 3000,
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
    const { email: _e, uf: _u, ...semEmail } = nova;
    res.status(201).json(semEmail);
  } catch (e) {
    console.error(e);
    res.status(500).json({ erro: "Não foi possível salvar sua participação agora." });
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

// Monta o WHERE a partir dos filtros da tela (região, estado, tipo e período)
function filtros(q) {
  const cond = [], valores = [];
  const add = (sql, v) => { valores.push(v); cond.push(sql.replace("?", `$${valores.length}`)); };

  if (q.uf === "nao-informado") cond.push("uf IS NULL");
  else if (UFS.includes(q.uf)) add("uf = ?", q.uf);
  else if (REGIOES[q.regiao]) add("uf = ANY(?)", REGIOES[q.regiao]);

  if (["conquista", "dedicatoria"].includes(q.tipo)) add("tipo = ?", q.tipo);
  if (/^\d{4}-\d{2}-\d{2}$/.test(q.de || "")) add("criado_em >= (?::date::timestamp AT TIME ZONE 'America/Sao_Paulo')", q.de);
  if (/^\d{4}-\d{2}-\d{2}$/.test(q.ate || "")) add("criado_em < ((?::date + 1)::timestamp AT TIME ZONE 'America/Sao_Paulo')", q.ate);
  if (q.exemplos !== "1") cond.push("id NOT LIKE 'exemplo-%'");

  return { where: cond.length ? `WHERE ${cond.join(" AND ")}` : "", valores };
}

app.get("/api/admin/metricas", exigirLogin, async (req, res) => {
  const { where, valores } = filtros(req.query);
  const hoje = `(now() AT TIME ZONE 'America/Sao_Paulo')::date`;
  const dia = `(criado_em AT TIME ZONE 'America/Sao_Paulo')::date`;
  try {
    const [resumo, porUf, porDia, musicas, artistas, ultimas] = await Promise.all([
      db.query(`SELECT count(*)::int AS total,
                       count(*) FILTER (WHERE tipo = 'conquista')::int AS conquistas,
                       count(*) FILTER (WHERE tipo = 'dedicatoria')::int AS dedicatorias,
                       count(*) FILTER (WHERE ${dia} = ${hoje})::int AS hoje,
                       count(*) FILTER (WHERE criado_em > now() - interval '1 hour')::int AS ultima_hora,
                       count(DISTINCT email)::int AS pessoas
                FROM participacoes ${where}`, valores),
      db.query(`SELECT uf, count(*)::int AS total FROM participacoes ${where} GROUP BY uf ORDER BY total DESC`, valores),
      db.query(`SELECT to_char(${dia}, 'YYYY-MM-DD') AS dia, count(*)::int AS total
                FROM participacoes ${where ? where + " AND" : "WHERE"} criado_em > now() - interval '30 days'
                GROUP BY 1 ORDER BY 1`, valores),
      db.query(`SELECT musica_nome AS nome, musica_artistas AS artistas, count(*)::int AS total
                FROM participacoes ${where} GROUP BY 1, 2 ORDER BY total DESC, nome LIMIT 10`, valores),
      db.query(`SELECT musica_artistas AS nome, count(*)::int AS total
                FROM participacoes ${where} GROUP BY 1 ORDER BY total DESC, nome LIMIT 10`, valores),
      db.query(`SELECT nome, email, tipo, amigo, uf, musica_nome, musica_artistas, criado_em
                FROM participacoes ${where} ORDER BY criado_em DESC LIMIT 20`, valores),
    ]);

    const regioes = Object.fromEntries(Object.keys(REGIOES).map((r) => [r, 0]));
    let semUf = 0;
    for (const { uf, total } of porUf.rows) {
      const r = regiaoDaUf(uf);
      if (r) regioes[r] += total; else semUf += total;
    }

    res.json({
      atualizadoEm: new Date().toISOString(),
      resumo: resumo.rows[0],
      porUf: porUf.rows.filter((r) => r.uf),
      porRegiao: { ...regioes, "Não informado": semUf },
      porDia: porDia.rows,
      musicas: musicas.rows,
      artistas: artistas.rows,
      ultimas: ultimas.rows,
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ erro: "Não foi possível carregar as métricas." });
  }
});

app.get("/api/admin/exportar", exigirLogin, async (req, res) => {
  const formato = req.query.formato === "xlsx" ? "xlsx" : "csv";
  const { where, valores } = filtros(req.query);
  try {
    const { rows } = await db.query(
      `SELECT * FROM participacoes ${where} ORDER BY criado_em DESC`, valores,
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