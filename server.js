const path = require("path");
const express = require("express");
const { google } = require("googleapis");

const app = express();
const PORT = process.env.PORT || 3000;

const SHEET_PUBLISH_ID =
  process.env.SHEET_PUBLISH_ID ||
  "2PACX-1vQL2uV2BS5DCGOlUQx4X2A7ABEWgC-c3CYA46B3S92pUG5H8VhFXta7qL00F3XjdqolkZ9jEPIqrp3Q";
const SHEET_TAB_NAME = process.env.SHEET_TAB_NAME || "Nomes";
const SHEET_CHAMADA_TAB_NAME = process.env.SHEET_CHAMADA_TAB_NAME || "Nomes Chamada";
const SHEET_CHAMADA_GID = process.env.SHEET_CHAMADA_GID || "934578770";
const GOOGLE_SPREADSHEET_ID =
  process.env.GOOGLE_SPREADSHEET_ID || "1spXWVi4VD1wIkGVXMVdcLpm9dHAgCJ5CTCBgmpiUji8";
const GOOGLE_SERVICE_ACCOUNT_EMAIL = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || "";
const GOOGLE_PRIVATE_KEY = process.env.GOOGLE_PRIVATE_KEY || "";

app.use(express.static("public"));
app.use(express.json());

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

function decodeHtml(text) {
  return text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function escapeRegex(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function extractTabGidFromPubHtml(html, tabName) {
  const safeTab = escapeRegex(tabName.trim());
  const regex = new RegExp(
    `href="([^"]*?[?&]gid=(\\d+)[^"]*?)"[^>]*>\\s*${safeTab}\\s*<`,
    "i"
  );
  const match = html.match(regex);
  return match ? match[2] : null;
}

function parseCsvFirstColumn(csvText) {
  return parseCsv(csvText)
    .map((r) => (r[0] || "").trim())
    .filter(Boolean);
}

function parseCsv(csvText) {
  const rows = [];
  let row = [];
  let field = "";
  let i = 0;
  let inQuotes = false;

  while (i < csvText.length) {
    const char = csvText[i];
    const next = csvText[i + 1];

    if (inQuotes) {
      if (char === '"' && next === '"') {
        field += '"';
        i += 2;
        continue;
      }
      if (char === '"') {
        inQuotes = false;
        i += 1;
        continue;
      }
      field += char;
      i += 1;
      continue;
    }

    if (char === '"') {
      inQuotes = true;
      i += 1;
      continue;
    }

    if (char === ",") {
      row.push(field);
      field = "";
      i += 1;
      continue;
    }

    if (char === "\r") {
      i += 1;
      continue;
    }

    if (char === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      i += 1;
      continue;
    }

    field += char;
    i += 1;
  }

  if (field.length > 0 || row.length > 0) {
    row.push(field);
    rows.push(row);
  }

  return rows;
}

function normalizeText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function toA1Column(colIndexZeroBased) {
  let n = colIndexZeroBased + 1;
  let result = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function isValidA1Cell(cell) {
  return /^[A-Z]+[1-9]\d*$/.test(String(cell || "").trim());
}

function getSheetsClient() {
  if (!GOOGLE_SERVICE_ACCOUNT_EMAIL || !GOOGLE_PRIVATE_KEY || !GOOGLE_SPREADSHEET_ID) {
    throw new Error(
      "Credenciais do Google Sheets nao configuradas. Configure GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_PRIVATE_KEY e GOOGLE_SPREADSHEET_ID."
    );
  }

  const auth = new google.auth.JWT(
    GOOGLE_SERVICE_ACCOUNT_EMAIL,
    null,
    GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),
    ["https://www.googleapis.com/auth/spreadsheets"]
  );

  return google.sheets({ version: "v4", auth });
}

function formatTodayPtBrShort() {
  const now = new Date();
  return `${now.getDate()}/${now.getMonth() + 1}`;
}

const TURMA_OPTIONS = [
  "8 Ano A",
  "8 Ano B",
  "9 Ano Anchieta",
  "1 Série Funcionários",
  "1 Série Anchieta",
  "2 Série Anchieta",
  "3 Série Anchieta"
];

const TURMA_ALIASES = {
  "8 ano a": ["8 a", "8 ano a"],
  "8 ano b": ["8 b", "8 ano b"],
  "9 ano anchieta": ["9 ano anchieta", "9 anchieta", "9 a"],
  "1 serie funcionarios": ["1 serie funcionarios", "1 serie", "1 série"],
  "1 serie anchieta": ["1 serie anchieta", "1 serie", "1 série"],
  "2 serie anchieta": ["2 serie anchieta", "2 serie", "2 série"],
  "3 serie anchieta": ["3 serie anchieta", "3 serie", "3 série"]
};

function findDateHeaderRow(rows) {
  return rows.findIndex((row) => normalizeText(row[0]) === "turma");
}

function extractDateColumns(rows) {
  const headerRowIndex = findDateHeaderRow(rows);
  if (headerRowIndex === -1) return { headerRowIndex: -1, dates: [] };

  const row = rows[headerRowIndex] || [];
  const datePattern = /^\d{1,2}\/\d{1,2}$/;
  const dates = [];

  for (let col = 0; col < row.length; col++) {
    const raw = String(row[col] || "").trim();
    if (!datePattern.test(raw)) continue;
    dates.push({
      value: raw,
      colIndex: col,
      a1Column: toA1Column(col)
    });
  }

  return { headerRowIndex, dates };
}

function findTurmaRow(rows, turmaSelection) {
  const normalizedSelection = normalizeText(turmaSelection);
  const aliases = TURMA_ALIASES[normalizedSelection] || [normalizedSelection];
  const aliasSet = new Set(aliases.map(normalizeText));

  for (let i = 0; i < rows.length; i++) {
    const cellA = normalizeText(rows[i]?.[0]);
    if (!cellA) continue;
    if (aliasSet.has(cellA)) {
      return i;
    }
  }

  return -1;
}

function isKnownTurmaLabel(cellA) {
  const normalized = normalizeText(cellA);
  if (!normalized) return false;

  const allAliases = Object.values(TURMA_ALIASES).flat().map(normalizeText);
  return allAliases.includes(normalized);
}

function extractStudentsForTurma(rows, turmaRowIndex, dateColIndex) {
  if (turmaRowIndex < 0) return [];
  const students = [];

  for (let i = turmaRowIndex + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const name = String(row[0] || "").trim();
    const rowHasAnyValue = row.some((cell) => String(cell || "").trim() !== "");

    if (!rowHasAnyValue) {
      if (students.length) break;
      continue;
    }

    if (!name) {
      if (students.length) break;
      continue;
    }

    if (isKnownTurmaLabel(name)) break;

    students.push({
      name,
      rowIndex: i,
      rowNumber: i + 1,
      currentValue:
        typeof dateColIndex === "number"
          ? String(row[dateColIndex] || "").trim().toUpperCase()
          : ""
    });
  }

  return students;
}

async function fetchChamadaSheetRows() {
  const csvUrl = `https://docs.google.com/spreadsheets/d/e/${SHEET_PUBLISH_ID}/pub?gid=${SHEET_CHAMADA_GID}&single=true&output=csv`;
  const csvResponse = await fetch(csvUrl, {
    headers: { "User-Agent": "Mozilla/5.0" }
  });

  if (!csvResponse.ok) {
    throw new Error(
      `Erro ao baixar CSV da aba ${SHEET_CHAMADA_TAB_NAME}: ${csvResponse.status}`
    );
  }

  const csvText = await csvResponse.text();
  return { rows: parseCsv(csvText), sourceUrl: csvUrl };
}

async function fetchChamadaData(selectedDate, selectedTurma) {
  const { rows, sourceUrl } = await fetchChamadaSheetRows();
  const { headerRowIndex, dates } = extractDateColumns(rows);

  const todaySuggestion = formatTodayPtBrShort();
  const chosenDate = String(selectedDate || todaySuggestion).trim();
  const chosenTurma = String(selectedTurma || TURMA_OPTIONS[0]).trim();

  const dateMatches = dates.filter((d) => d.value === chosenDate);
  const selectedDateColumn = dateMatches[0] || null;

  const turmaRowIndex = findTurmaRow(rows, chosenTurma);
  const turmaRowNumber = turmaRowIndex >= 0 ? turmaRowIndex + 1 : null;
  const turmaCell = turmaRowNumber ? `A${turmaRowNumber}` : null;
  const students = extractStudentsForTurma(
    rows,
    turmaRowIndex,
    selectedDateColumn ? selectedDateColumn.colIndex : undefined
  );

  return {
    sourceUrl,
    sheet: SHEET_CHAMADA_TAB_NAME,
    todaySuggestion,
    turmaOptions: TURMA_OPTIONS,
    availableDates: dates.map((d) => d.value),
    selected: {
      date: chosenDate,
      turma: chosenTurma,
      dateColumn: selectedDateColumn,
      dateMatchesCount: dateMatches.length,
      dateHeaderRow: headerRowIndex >= 0 ? headerRowIndex + 1 : null,
      turmaRow: turmaRowNumber,
      turmaCell
    },
    students
  };
}

async function fetchNamesFromPublishedCsv() {
  const pubHtmlUrl = `https://docs.google.com/spreadsheets/d/e/${SHEET_PUBLISH_ID}/pubhtml`;
  const pubHtmlResponse = await fetch(pubHtmlUrl, {
    headers: { "User-Agent": "Mozilla/5.0" }
  });

  if (!pubHtmlResponse.ok) {
    throw new Error(
      `Erro ao consultar página publicada da planilha: ${pubHtmlResponse.status}`
    );
  }

  const pubHtml = decodeHtml(await pubHtmlResponse.text());
  const gid = extractTabGidFromPubHtml(pubHtml, SHEET_TAB_NAME);

  if (!gid) {
    throw new Error(`Não foi possível localizar a aba "${SHEET_TAB_NAME}".`);
  }

  const csvUrl = `https://docs.google.com/spreadsheets/d/e/${SHEET_PUBLISH_ID}/pub?gid=${gid}&single=true&output=csv`;
  const csvResponse = await fetch(csvUrl, {
    headers: { "User-Agent": "Mozilla/5.0" }
  });

  if (!csvResponse.ok) {
    throw new Error(`Erro ao baixar CSV da aba ${SHEET_TAB_NAME}: ${csvResponse.status}`);
  }

  const csvText = await csvResponse.text();
  const names = parseCsvFirstColumn(csvText);

  return {
    names,
    sourceUrl: csvUrl,
    sheet: SHEET_TAB_NAME,
    updatedAt: new Date().toISOString()
  };
}

async function fetchNames() {
  const url = `https://docs.google.com/spreadsheets/d/e/${SHEET_PUBLISH_ID}/gviz/tq?sheet=${encodeURIComponent(
    SHEET_TAB_NAME
  )}&tqx=out:json`;

  try {
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
  } catch (_error) {
    // Fallback para planilhas publicadas no formato /d/e/.../pubhtml
    return fetchNamesFromPublishedCsv();
  }
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

app.get("/api/chamada", async (req, res) => {
  try {
    const data = await fetchChamadaData(req.query.date, req.query.turma);
    res.json(data);
  } catch (error) {
    res.status(500).json({
      error: "Falha ao carregar dados da aba Nomes Chamada.",
      details: error.message
    });
  }
});

app.post("/api/chamada/marcar", async (req, res) => {
  try {
    const cell = String(req.body?.cell || "")
      .trim()
      .toUpperCase();
    const value = String(req.body?.value || "")
      .trim()
      .toUpperCase();

    if (!isValidA1Cell(cell)) {
      return res.status(400).json({
        error: "Celula invalida. Use formato A1 (ex.: F23)."
      });
    }

    if (value !== "" && value !== "F") {
      return res.status(400).json({
        error: "Valor invalido. Use 'F' para marcar falta ou vazio para limpar."
      });
    }

    const sheets = getSheetsClient();
    const range = `'${SHEET_CHAMADA_TAB_NAME}'!${cell}`;

    await sheets.spreadsheets.values.update({
      spreadsheetId: GOOGLE_SPREADSHEET_ID,
      range,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[value]]
      }
    });

    res.json({
      ok: true,
      sheet: SHEET_CHAMADA_TAB_NAME,
      cell,
      value
    });
  } catch (error) {
    res.status(500).json({
      error: "Falha ao gravar no Google Sheets.",
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
