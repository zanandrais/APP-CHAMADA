# Webservice de Nomes (Google Sheets + Render)

Este projeto exibe em tela os valores da primeira coluna da aba `Nomes` de uma planilha Google Sheets publicada.

## Como rodar localmente

```bash
npm install
npm start
```

Abra `http://localhost:3000`.

## Deploy no Render

1. Suba estes arquivos para um repositório Git (GitHub/GitLab).
2. No Render, crie um `Web Service` usando o repositório.
3. O Render pode usar o arquivo `render.yaml` automaticamente.
4. Se configurar manualmente:
   - Build Command: `npm install`
   - Start Command: `npm start`
5. Variáveis de ambiente (opcional):
   - `SHEET_PUBLISH_ID`
   - `SHEET_TAB_NAME` (padrão: `Nomes`)

## Endpoints

- `/` interface HTML/CSS
- `/api/nomes` retorna JSON com os nomes
