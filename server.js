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
  "9 ano anchieta": ["9 ano anchieta", "9 anchieta", "9 ano", "9 a", "9"],
  "1 serie funcionarios": ["1 serie funcionarios", "1 serie", "1 série"],
  "1 serie anchieta": ["1 serie anchieta", "1 serie", "1 série"],
  "2 serie anchieta": ["2 serie anchieta", "2 serie", "2 série"],
  "3 serie anchieta": ["3 serie anchieta", "3 serie", "3 série"]
};

function parseTurmaDescriptor(text) {
  const normalized = normalizeText(text);
  if (!normalized) return null;

  const tokens = normalized.split(" ").filter(Boolean);
  const numberMatch = normalized.match(/^(\d{1,2})/);

  return {
    normalized,
    tokens,
    number: numberMatch ? numberMatch[1] : null,
    letter: tokens.find((token) => /^[ab]$/.test(token)) || null,
    stage: tokens.includes("ano") ? "ano" : tokens.includes("serie") ? "serie" : null,
    campus: tokens.includes("anchieta")
      ? "anchieta"
      : tokens.includes("funcionarios")
        ? "funcionarios"
        : null
  };
}

function looksLikeTurmaHeader(cellA) {
  const descriptor = parseTurmaDescriptor(cellA);
  if (!descriptor?.number) return false;
  if (!/^\d{1,2}/.test(descriptor.normalized)) return false;

  // Turma labels are short markers, unlike student full names.
  if (descriptor.tokens.length > 4) return false;

  return Boolean(descriptor.stage || descriptor.letter || descriptor.campus);
}

function scoreTurmaCandidate(selectedDescriptor, candidateDescriptor) {
  if (!selectedDescriptor || !candidateDescriptor?.number) {
    return Number.NEGATIVE_INFINITY;
  }

  let score = 0;

  if (selectedDescriptor.number) {
    if (selectedDescriptor.number !== candidateDescriptor.number) {
      return Number.NEGATIVE_INFINITY;
    }
    score += 100;
  }

  if (selectedDescriptor.letter) {
    if (candidateDescriptor.letter && selectedDescriptor.letter !== candidateDescriptor.letter) {
      return Number.NEGATIVE_INFINITY;
    }
    score += candidateDescriptor.letter ? 20 : 4;
  }

  if (selectedDescriptor.stage) {
    if (candidateDescriptor.stage === selectedDescriptor.stage) {
      score += 12;
    } else if (!candidateDescriptor.stage) {
      score += 2;
    }
  }

  if (selectedDescriptor.campus) {
    if (candidateDescriptor.campus && candidateDescriptor.campus !== selectedDescriptor.campus) {
      return Number.NEGATIVE_INFINITY;
    }
    score += candidateDescriptor.campus ? 18 : 1;
  }

  if (candidateDescriptor.normalized === selectedDescriptor.normalized) {
    score += 30;
  }

  return score;
}

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
  const selectedDescriptor = parseTurmaDescriptor(turmaSelection);
  const candidates = [];

  for (let i = 0; i < rows.length; i++) {
    const rawCellA = String(rows[i]?.[0] || "").trim();
    const normalizedCellA = normalizeText(rawCellA);
    if (!normalizedCellA) continue;

    const aliasMatch = aliasSet.has(normalizedCellA);
    const headerLike = looksLikeTurmaHeader(rawCellA);
    if (!aliasMatch && !headerLike) continue;

    const candidateDescriptor = parseTurmaDescriptor(rawCellA);
    const descriptorScore = scoreTurmaCandidate(selectedDescriptor, candidateDescriptor);

    if (descriptorScore === Number.NEGATIVE_INFINITY && !aliasMatch) {
      continue;
    }

    let score = aliasMatch ? 200 : 0;
    if (descriptorScore !== Number.NEGATIVE_INFINITY) {
      score += descriptorScore;
    }

    candidates.push({
      rowIndex: i,
      score
    });
  }

  if (!candidates.length) return -1;

  candidates.sort((a, b) => b.score - a.score || a.rowIndex - b.rowIndex);

  const bestScore = candidates[0].score;
  const bestCandidates = candidates.filter((candidate) => candidate.score === bestScore);

  if (
    bestCandidates.length > 1 &&
    selectedDescriptor?.number === "1" &&
    selectedDescriptor?.stage === "serie" &&
    selectedDescriptor?.campus
  ) {
    if (selectedDescriptor.campus === "funcionarios") {
      return bestCandidates[0].rowIndex;
    }

    if (selectedDescriptor.campus === "anchieta") {
      return bestCandidates[bestCandidates.length - 1].rowIndex;
    }
  }

  return bestCandidates[0].rowIndex;
}

function isKnownTurmaLabel(cellA) {
  if (looksLikeTurmaHeader(cellA)) return true;

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

function buildTurmaSelectionData(rows, sheetName, selectedDate, selectedTurma) {
  const { headerRowIndex, dates } = extractDateColumns(rows);
  const dateMatches = dates.filter((d) => d.value === selectedDate);
  const selectedDateColumn = dateMatches[0] || null;

  const turmaRowIndex = findTurmaRow(rows, selectedTurma);
  const turmaRowNumber = turmaRowIndex >= 0 ? turmaRowIndex + 1 : null;
  const turmaCell = turmaRowNumber ? `A${turmaRowNumber}` : null;
  const students = extractStudentsForTurma(
    rows,
    turmaRowIndex,
    selectedDateColumn ? selectedDateColumn.colIndex : undefined
  );

  return {
    sheet: sheetName,
    availableDates: dates.map((d) => d.value),
    selected: {
      date: selectedDate,
      turma: selectedTurma,
      dateColumn: selectedDateColumn,
      dateMatchesCount: dateMatches.length,
      dateHeaderRow: headerRowIndex >= 0 ? headerRowIndex + 1 : null,
      turmaRow: turmaRowNumber,
      turmaCell
    },
    students
  };
}

function buildStudentNameBuckets(students) {
  const buckets = new Map();

  for (const student of students || []) {
    const key = normalizeText(student.name);
    if (!key) continue;
    if (!buckets.has(key)) buckets.set(key, []);
    buckets.get(key).push(student);
  }

  return buckets;
}

function mergeStudentsWithNomes(chamadaStudents, nomesSelection) {
  const nomesStudents = nomesSelection?.students || [];
  const namesBuckets = buildStudentNameBuckets(nomesStudents);

  return (chamadaStudents || []).map((student, index) => {
    const nameKey = normalizeText(student.name);
    let nomesStudent = null;

    if (nameKey && namesBuckets.has(nameKey)) {
      const bucket = namesBuckets.get(nameKey);
      nomesStudent = bucket.shift() || null;
    }

    // Fallback por posição relativa quando os nomes divergem levemente entre abas.
    if (!nomesStudent && index < nomesStudents.length) {
      nomesStudent = nomesStudents[index];
    }

    const nomesDateColumn = nomesSelection?.selected?.dateColumn || null;
    const nomesCell =
      nomesStudent && nomesDateColumn
        ? `${nomesDateColumn.a1Column}${nomesStudent.rowNumber}`
        : null;

    return {
      ...student,
      nomesRowNumber: nomesStudent ? nomesStudent.rowNumber : null,
      nomesCurrentValue: nomesStudent ? nomesStudent.currentValue || "" : "",
      nomesCell,
      nomesMatched: Boolean(nomesStudent)
    };
  });
}

async function fetchPublishedTabRowsByGid(tabName, gid) {
  const csvUrl = `https://docs.google.com/spreadsheets/d/e/${SHEET_PUBLISH_ID}/pub?gid=${gid}&single=true&output=csv`;
  const csvResponse = await fetch(csvUrl, {
    headers: { "User-Agent": "Mozilla/5.0" }
  });

  if (!csvResponse.ok) {
    throw new Error(`Erro ao baixar CSV da aba ${tabName}: ${csvResponse.status}`);
  }

  const csvText = await csvResponse.text();
  return { rows: parseCsv(csvText), sourceUrl: csvUrl, gid };
}

async function fetchPublishedTabRowsByName(tabName) {
  const pubHtmlUrl = `https://docs.google.com/spreadsheets/d/e/${SHEET_PUBLISH_ID}/pubhtml`;
  const pubHtmlResponse = await fetch(pubHtmlUrl, {
    headers: { "User-Agent": "Mozilla/5.0" }
  });

  if (!pubHtmlResponse.ok) {
    throw new Error(
      `Erro ao consultar pagina publicada da planilha: ${pubHtmlResponse.status}`
    );
  }

  const pubHtml = decodeHtml(await pubHtmlResponse.text());
  const gid = extractTabGidFromPubHtml(pubHtml, tabName);

  if (!gid) {
    throw new Error(`Nao foi possivel localizar a aba "${tabName}".`);
  }

  return fetchPublishedTabRowsByGid(tabName, gid);
}

async function fetchChamadaSheetRows() {
  return fetchPublishedTabRowsByGid(SHEET_CHAMADA_TAB_NAME, SHEET_CHAMADA_GID);
}

async function fetchChamadaData(selectedDate, selectedTurma) {
  const todaySuggestion = formatTodayPtBrShort();
  const chosenDate = String(selectedDate || todaySuggestion).trim();
  const chosenTurma = String(selectedTurma || TURMA_OPTIONS[0]).trim();
  const [chamadaSource, nomesSource] = await Promise.all([
    fetchChamadaSheetRows(),
    fetchPublishedTabRowsByName(SHEET_TAB_NAME)
  ]);

  const chamadaSelection = buildTurmaSelectionData(
    chamadaSource.rows,
    SHEET_CHAMADA_TAB_NAME,
    chosenDate,
    chosenTurma
  );
  const nomesSelection = buildTurmaSelectionData(
    nomesSource.rows,
    SHEET_TAB_NAME,
    chosenDate,
    chosenTurma
  );
  const mergedStudents = mergeStudentsWithNomes(chamadaSelection.students, nomesSelection);

  return {
    sourceUrl: chamadaSource.sourceUrl,
    sourceUrlNomes: nomesSource.sourceUrl,
    sheet: SHEET_CHAMADA_TAB_NAME,
    sheetNomes: SHEET_TAB_NAME,
    todaySuggestion,
    turmaOptions: TURMA_OPTIONS,
    availableDates: chamadaSelection.availableDates,
    selected: chamadaSelection.selected,
    nomesSelection: {
      selected: nomesSelection.selected,
      availableDates: nomesSelection.availableDates
    },
    students: mergedStudents
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
    const sheet = String(req.body?.sheet || SHEET_CHAMADA_TAB_NAME).trim();
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

    const allowedSheets = new Set([SHEET_CHAMADA_TAB_NAME, SHEET_TAB_NAME]);
    if (!allowedSheets.has(sheet)) {
      return res.status(400).json({
        error: "Aba invalida para gravacao."
      });
    }

    if (!["", "F", "1", "2"].includes(value)) {
      return res.status(400).json({
        error: "Valor invalido. Use 'F', '1', '2' ou vazio para limpar."
      });
    }

    if (sheet === SHEET_CHAMADA_TAB_NAME && value !== "" && value !== "F") {
      return res.status(400).json({
        error: "Na aba Nomes Chamada, use apenas 'F' ou vazio."
      });
    }

    if (sheet === SHEET_TAB_NAME && value !== "" && value !== "1" && value !== "2") {
      return res.status(400).json({
        error: "Na aba Nomes, use apenas '1', '2' ou vazio."
      });
    }

    const sheets = getSheetsClient();
    const range = `'${sheet}'!${cell}`;

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
      sheet,
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
