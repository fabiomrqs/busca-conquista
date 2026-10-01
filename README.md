# Qual é a música da sua conquista? · Senac EAD

Fluxo de 5 telas: início → dados → busca no Spotify → cartão de agradecimento → mural.

## Rodar

1. Crie um app em https://developer.spotify.com/dashboard e copie o Client ID e o Client Secret.
2. Copie `.env.example` para `.env` e preencha as chaves.
3. Instale e rode (Node 20.6+):

       npm install
       npm start

4. Abra http://localhost:3000

## Estrutura

    server.js            busca no Spotify + API do mural
    public/index.html    as 5 telas (HTML, CSS e JS num arquivo só)
    public/img/img-1.png logo "Quer saber? Senac EAD!" (substitua)
    public/img/img-2.png logo Senac (substitua)
    data/                criado automaticamente com as participações

## Ajustes

- Link da playlist e URL de compartilhamento: objeto `CONFIG` no início do `<script>` em `public/index.html`.
- Link dos termos de privacidade: `href="#"` na tela 2.

## API

- `GET /api/buscar?q=termo` → faixas do Spotify.
- `GET /api/participacoes` → cards do mural (sem e-mail).
- `POST /api/participacoes` → `{ nome, email, tipo: "conquista"|"dedicatoria", amigo?, aceite: true, musicaId }`.
  O servidor confirma a faixa no Spotify antes de salvar.

As participações ficam em `data/participacoes.json` (começa com 6 exemplos).
Para produção, troque `lerTodas`/`salvarTodas` no `server.js` por um banco de dados.
