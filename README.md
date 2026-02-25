# Chamada (Google Sheets + Render)

App web para:
- selecionar `data` e `turma`
- localizar a coluna/linha na aba `Nomes Chamada`
- listar alunos
- marcar/desmarcar `F` (falta) diretamente no Google Sheets

## Rodar localmente

```bash
npm install
npm start
```

Acesse `http://localhost:3000`.

## Deploy no Render

Use o `render.yaml` ou configure manualmente:
- Build Command: `npm install`
- Start Command: `npm start`

## Variáveis de ambiente (leitura)

- `SHEET_PUBLISH_ID`
- `SHEET_CHAMADA_TAB_NAME` (padrão: `Nomes Chamada`)
- `SHEET_CHAMADA_GID` (aba `Nomes Chamada`)

## Variáveis para gravação no Google Sheets

- `GOOGLE_SPREADSHEET_ID`
- `GOOGLE_SERVICE_ACCOUNT_EMAIL`
- `GOOGLE_PRIVATE_KEY`

## Como habilitar a gravação (botão F)

1. No Google Cloud, crie um projeto.
2. Ative a `Google Sheets API`.
3. Crie uma `Service Account`.
4. Gere uma chave JSON.
5. Compartilhe a planilha com o e-mail da Service Account como `Editor`.
6. No Render, adicione:
   - `GOOGLE_SERVICE_ACCOUNT_EMAIL` (campo `client_email` do JSON)
   - `GOOGLE_PRIVATE_KEY` (campo `private_key` do JSON)

Observação:
- Se colar a chave no Render com `\n`, o backend já converte automaticamente.

## Endpoints

- `GET /` interface de chamada
- `GET /api/chamada` leitura de data/turma/alunos
- `POST /api/chamada/marcar` grava `F` ou limpa uma célula (`{ "cell": "F23", "value": "F" }`)
