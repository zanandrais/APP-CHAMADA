const path = require("path");
const express = require("express");

const app = express();
const PORT = process.env.PORT || 3000;

const SHEET_PUBLISH_ID =
  process.env.SHEET_PUBLISH_ID ||
  "2PACX-1vQL2uV2BS5DCGOlUQx4X2A7ABEWgC-c3CYA46B3S92pUG5H8VhFXta7qL00F3XjdqolkZ9jEPIqrp3Q";
const SHEET_TAB_NAME = process.env.SHEET_TAB_NAME || "Nomes";

app.use(express.static("public"));

function parseGoogleVizResponse(text) {
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");

  if (start === -1 || end === -1 || end <= start) {
    throw new Error("Resposta gviz em formato inesperado.");
  }

  return JSON.parse(text.slice(start, end + 1));
}

function normalizeCellValue(cell) {
  if (!cell) return "";

  if (typeof cell.f === "string" && cell.f.trim()) {
    return cell.f.trim();
  }

  if (cell.v == null) return "";
  return String(cell.v).trim();
}

function extractNamesFromTable(table) {
  if (!table || !Array.isArray(table.rows)) {
    return [];
  }

  const names = [];

  for (const row of table.rows) {
    if (!row || !Array.isArray(row.c) || row.c.length === 0) continue;

    const firstColumn = normalizeCellValue(row.c[0]);
    if (!firstColumn) continue;

    names.push(firstColumn);
  }

  return names;
}

async function fetchNames() {
  const url = `https://docs.google.com/spreadsheets/d/e/${SHEET_PUBLISH_ID}/gviz/tq?sheet=${encodeURIComponent(
    SHEET_TAB_NAME
  )}&tqx=out:json`;

  const response = await fetch(url, {
    headers: {
      "User-Agent": "Mozilla/5.0"
    }
  });

  if (!response.ok) {
    throw new Error(`Erro ao consultar Google Sheets: ${response.status}`);
  }

  const text = await response.text();
  const payload = parseGoogleVizResponse(text);
  const names = extractNamesFromTable(payload.table);

  return {
    names,
    sourceUrl: url,
    sheet: SHEET_TAB_NAME,
    updatedAt: new Date().toISOString()
  };
}

app.get("/api/nomes", async (_req, res) => {
  try {
    const data = await fetchNames();
    res.json(data);
  } catch (error) {
    res.status(500).json({
      error: "Falha ao carregar nomes da planilha.",
      details: error.message
    });
  }
});

app.get("/", (_req, res) => {
  res.sendFile(path.join(process.cwd(), "public", "index.html"));
});

app.listen(PORT, () => {
  console.log(`Servidor rodando na porta ${PORT}`);
});
