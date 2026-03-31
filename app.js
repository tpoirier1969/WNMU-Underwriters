const STORAGE_KEY = 'wnmu-underwriter-intake-v0.1.0';
const OCR_PAGE_LIMIT = 4;

const state = {
  records: [],
  selectedId: null,
  importQueueActive: false,
};

const els = {
  dropZone: document.getElementById('dropZone'),
  fileInput: document.getElementById('fileInput'),
  importStatus: document.getElementById('importStatus'),
  recordsBody: document.getElementById('recordsBody'),
  recordForm: document.getElementById('recordForm'),
  detailBadge: document.getElementById('detailBadge'),
  recordIssues: document.getElementById('recordIssues'),
  searchInput: document.getElementById('searchInput'),
  issueFilter: document.getElementById('issueFilter'),
  metrics: document.getElementById('metrics'),
  quarterStart: document.getElementById('quarterStart'),
  quarterEnd: document.getElementById('quarterEnd'),
  narrativeBox: document.getElementById('narrativeBox'),
  rowTemplate: document.getElementById('rowTemplate'),
  importJsonInput: document.getElementById('importJsonInput'),
};

boot();

function boot() {
  loadState();
  bindEvents();
  recalcFlags();
  renderAll();
}

function bindEvents() {
  els.fileInput.addEventListener('change', async (event) => {
    await importFiles([...event.target.files]);
    event.target.value = '';
  });

  ['dragenter', 'dragover'].forEach((evt) => {
    els.dropZone.addEventListener(evt, (event) => {
      event.preventDefault();
      event.stopPropagation();
      els.dropZone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach((evt) => {
    els.dropZone.addEventListener(evt, (event) => {
      event.preventDefault();
      event.stopPropagation();
      els.dropZone.classList.remove('dragover');
    });
  });

  els.dropZone.addEventListener('drop', async (event) => {
    const files = [...event.dataTransfer.files].filter(Boolean);
    await importFiles(files);
  });

  els.dropZone.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      els.fileInput.click();
    }
  });

  document.getElementById('manualAddBtn').addEventListener('click', () => {
    const record = makeBlankRecord();
    state.records.unshift(record);
    state.selectedId = record.id;
    persist();
    recalcFlags();
    renderAll();
  });

  document.getElementById('saveRecordBtn').addEventListener('click', () => saveFormToSelected());
  document.getElementById('deleteRecordBtn').addEventListener('click', () => deleteSelected());
  document.getElementById('duplicateRecordBtn').addEventListener('click', () => duplicateSelected());
  document.getElementById('clearAllBtn').addEventListener('click', clearAllRecords);
  document.getElementById('exportQuarterCsvBtn').addEventListener('click', exportQuarterCsv);
  document.getElementById('copyNarrativeBtn').addEventListener('click', copyNarrative);
  document.getElementById('exportJsonBtn').addEventListener('click', exportJsonBackup);
  document.getElementById('importJsonBtn').addEventListener('click', () => els.importJsonInput.click());
  els.importJsonInput.addEventListener('change', importJsonBackup);

  [els.searchInput, els.issueFilter, els.quarterStart, els.quarterEnd].forEach((el) => {
    el.addEventListener('input', () => {
      recalcFlags();
      renderAll();
    });
    el.addEventListener('change', () => {
      recalcFlags();
      renderAll();
    });
  });
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed.records)) {
      state.records = parsed.records.map(sanitizeRecord);
      state.selectedId = parsed.selectedId || state.records[0]?.id || null;
    }
  } catch (error) {
    console.error('Failed to load local state', error);
  }
}

function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({
    records: state.records,
    selectedId: state.selectedId,
  }));
}

function makeBlankRecord() {
  return sanitizeRecord({
    id: crypto.randomUUID(),
    underwriterName: '',
    contractType: '',
    programName: '',
    contactPerson: '',
    email: '',
    phone: '',
    startDate: '',
    endDate: '',
    amount: '',
    creditCount: '',
    programCount: '',
    creditCopy: '',
    notes: '',
    rawText: '',
    sourceFileName: 'Manual entry',
    sourceHash: '',
    importedAt: new Date().toISOString(),
    issues: [],
  });
}

function sanitizeRecord(record) {
  const blank = makeTrulyBlank();
  return { ...blank, ...record, issues: Array.isArray(record?.issues) ? record.issues : [] };
}

function makeTrulyBlank() {
  return {
    id: crypto.randomUUID(),
    underwriterName: '',
    contractType: '',
    programName: '',
    contactPerson: '',
    email: '',
    phone: '',
    startDate: '',
    endDate: '',
    amount: '',
    creditCount: '',
    programCount: '',
    creditCopy: '',
    notes: '',
    rawText: '',
    sourceFileName: '',
    sourceHash: '',
    importedAt: new Date().toISOString(),
    issues: [],
  };
}

async function importFiles(files) {
  if (!files.length) return;
  if (state.importQueueActive) {
    setStatus('Import already running. Finish that first so the browser does not catch fire.');
    return;
  }
  state.importQueueActive = true;

  try {
    const accepted = files.filter((file) => /\.(pdf|docx?|json)$/i.test(file.name));
    if (!accepted.length) {
      setStatus('No supported files found. Use PDF, DOCX, or JSON backup.');
      return;
    }

    for (let i = 0; i < accepted.length; i += 1) {
      const file = accepted[i];
      setStatus(`Processing ${i + 1} of ${accepted.length}: ${file.name}`);
      if (/\.json$/i.test(file.name)) {
        await importJsonFile(file);
        continue;
      }
      const record = await parseFileToRecord(file, i + 1, accepted.length);
      upsertRecord(record);
      state.selectedId = record.id;
      recalcFlags();
      renderAll();
      persist();
    }
    setStatus(`Finished importing ${accepted.length} file(s).`);
  } catch (error) {
    console.error(error);
    setStatus(`Import failed: ${error.message || error}`);
  } finally {
    state.importQueueActive = false;
    recalcFlags();
    renderAll();
    persist();
  }
}

async function parseFileToRecord(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const hash = await hashBytes(bytes);
  let rawText = '';
  let extractionPath = '';

  if (/\.docx?$/i.test(file.name)) {
    rawText = await extractDocxText(bytes.buffer);
    extractionPath = 'DOCX text extraction';
  } else if (/\.pdf$/i.test(file.name)) {
    const pdfResult = await extractPdfText(bytes);
    rawText = pdfResult.text;
    extractionPath = pdfResult.path;
  } else {
    rawText = '';
    extractionPath = 'Unsupported file type';
  }

  const parsed = parseContractText(rawText, file.name);
  return sanitizeRecord({
    ...parsed,
    id: crypto.randomUUID(),
    sourceFileName: file.name,
    sourceHash: hash,
    importedAt: new Date().toISOString(),
    notes: parsed.notes || `Imported via ${extractionPath}`,
    rawText,
  });
}

async function extractDocxText(arrayBuffer) {
  const result = await mammoth.extractRawText({ arrayBuffer });
  return normalizeText(result.value || '');
}

async function extractPdfText(bytes) {
  const pdfjsLib = await import('https://cdn.jsdelivr.net/npm/pdfjs-dist@5.6.205/build/pdf.min.mjs');
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@5.6.205/build/pdf.worker.min.mjs';

  const loadingTask = pdfjsLib.getDocument({ data: bytes });
  const pdf = await loadingTask.promise;
  let textParts = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map((item) => item.str).join(' ').trim();
    if (pageText.length > 80) {
      textParts.push(pageText);
    }
  }

  const directText = normalizeText(textParts.join('\n\n'));
  if (directText.length > 120) {
    return { text: directText, path: 'PDF text layer' };
  }

  const ocrParts = [];
  const pageLimit = Math.min(pdf.numPages, OCR_PAGE_LIMIT);
  for (let pageNumber = 1; pageNumber <= pageLimit; pageNumber += 1) {
    setStatus(`Running OCR on ${pageNumber}/${pageLimit} page(s)...`);
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale: 2 });
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({ canvasContext: context, viewport }).promise;
    const result = await Tesseract.recognize(canvas, 'eng', {
      logger: () => {},
    });
    ocrParts.push(result.data?.text || '');
  }

  return { text: normalizeText(ocrParts.join('\n\n')), path: `PDF OCR (${pageLimit} page max)` };
}

function parseContractText(rawText, fileName = '') {
  const text = normalizeText(rawText);
  const lines = text.split('\n').map((line) => line.trim()).filter(Boolean);
  const lower = text.toLowerCase();

  const dates = extractDates(text).sort();
  const moneyValues = extractCurrency(text);
  const emails = uniqueMatches(text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || []);
  const phones = uniqueMatches(text.match(/(?:\+?1[-.\s]?)?(?:\(?\d{3}\)?[-.\s]?)\d{3}[-.\s]?\d{4}/g) || []);

  const record = makeTrulyBlank();
  record.underwriterName = findNamedField(text, [
    /(?:underwriter|sponsor|organization|company|business(?:\s+name)?|client)\s*[:\-]\s*(.+)/i,
    /(?:this agreement is between|agreement between|contract between)\s+(.+?)\s+(?:and|for)\s+/i,
  ]) || guessEntityLine(lines, fileName);

  record.contractType = detectContractType(lower, lines, fileName);
  record.programName = findNamedField(text, [
    /(?:program|series|campaign|show|special)\s*[:\-]\s*(.+)/i,
    /(?:title of program|contract title)\s*[:\-]\s*(.+)/i,
  ]) || guessProgramName(lines, lower);

  record.contactPerson = findNamedField(text, [
    /(?:contact person|contact|representative|sales rep|account executive)\s*[:\-]\s*(.+)/i,
  ]);
  record.email = emails[0] || '';
  record.phone = phones[0] || '';

  const explicitStart = findNamedField(text, [
    /(?:start date|begins?|effective date|contract start)\s*[:\-]\s*([A-Za-z]+\s+\d{1,2},\s*\d{4}|\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i,
  ]);
  const explicitEnd = findNamedField(text, [
    /(?:end date|ends?|expiration date|contract end)\s*[:\-]\s*([A-Za-z]+\s+\d{1,2},\s*\d{4}|\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i,
  ]);

  record.startDate = normalizeDate(explicitStart || dates[0] || '');
  record.endDate = normalizeDate(explicitEnd || dates[dates.length - 1] || '');

  const creditCountMatch = text.match(/(\d{1,3})\s+(?:credits?|spots?|announcements?)/i);
  const programCountMatch = text.match(/(\d{1,3})\s+(?:programs?|episodes?|airings?)/i);
  record.creditCount = creditCountMatch?.[1] || '';
  record.programCount = programCountMatch?.[1] || '';

  const amount = moneyValues.length ? Math.max(...moneyValues) : '';
  record.amount = amount ? String(amount) : '';

  record.creditCopy = findNamedField(text, [
    /(?:credit copy|copy|announcement|audio credit|underwriting copy)\s*[:\-]\s*([\s\S]{20,400})/i,
  ], true) || extractQuotedBlock(text);

  if (!record.underwriterName && record.email) {
    record.underwriterName = emailDomainToName(record.email);
  }

  return record;
}

function detectContractType(lower, lines, fileName) {
  if (lower.includes('american revolution')) return 'American Revolution';
  if (lower.includes('underwriter') && lower.includes('contract')) return 'Underwriter Contract';
  if (lower.includes('underwriting')) return 'Underwriting';
  if (fileName) return fileName.replace(/\.[^.]+$/, '');
  return '';
}

function guessProgramName(lines, lower) {
  const candidates = lines.filter((line) => {
    const l = line.toLowerCase();
    return l.length > 5 && l.length < 80 && !/(wnmu|contract|agreement|phone|email|address|date|amount|signature)/i.test(l);
  });
  const known = candidates.find((line) => /american revolution|pbs|program|series|special/i.test(line));
  return known || '';
}

function guessEntityLine(lines, fileName) {
  const filtered = lines.filter((line) => {
    if (line.length < 3 || line.length > 80) return false;
    if (/(wnmu|western|public broadcasting|contract|agreement|phone|email|address|date|start|end|amount|signature|copy)/i.test(line)) return false;
    return true;
  });
  const titleCaseLine = filtered.find((line) => /[A-Z][a-z]+\s+[A-Z]/.test(line));
  if (titleCaseLine) return titleCaseLine;
  return fileName.replace(/\.[^.]+$/, '');
}

function emailDomainToName(email) {
  const domain = (email.split('@')[1] || '').replace(/\.[a-z]{2,}$/i, '');
  return domain
    .split(/[.-]/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(' ');
}

function extractQuotedBlock(text) {
  const match = text.match(/["“]([^"”]{20,500})["”]/);
  return match?.[1]?.trim() || '';
}

function findNamedField(text, patterns, multiLine = false) {
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match?.[1]) {
      const value = multiLine ? match[1].trim() : match[1].split('\n')[0].trim();
      return value.replace(/\s+/g, ' ').trim();
    }
  }
  return '';
}

function extractDates(text) {
  const matches = [
    ...(text.match(/\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},\s*\d{4}\b/gi) || []),
    ...(text.match(/\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b/g) || []),
  ];
  return uniqueMatches(matches).filter((date) => normalizeDate(date));
}

function extractCurrency(text) {
  const matches = text.match(/\$\s*\d[\d,]*(?:\.\d{2})?/g) || [];
  return uniqueMatches(matches)
    .map((value) => Number(value.replace(/[^\d.]/g, '')))
    .filter((value) => Number.isFinite(value));
}

function normalizeDate(input) {
  if (!input) return '';
  const trimmed = String(input).trim();
  const numeric = trimmed.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  let date;
  if (numeric) {
    let [, mm, dd, yy] = numeric;
    let year = Number(yy);
    if (year < 100) year += year >= 70 ? 1900 : 2000;
    date = new Date(Date.UTC(year, Number(mm) - 1, Number(dd)));
  } else {
    date = new Date(trimmed);
  }
  if (Number.isNaN(date.getTime())) return '';
  return date.toISOString().slice(0, 10);
}

function normalizeText(text) {
  return String(text || '')
    .replace(/\r/g, '\n')
    .replace(/\t/g, ' ')
    .replace(/[ \u00A0]{2,}/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function uniqueMatches(values) {
  return [...new Set(values.map((v) => String(v).trim()).filter(Boolean))];
}

function upsertRecord(record) {
  const existingIndex = state.records.findIndex((item) => item.sourceHash && item.sourceHash === record.sourceHash);
  if (existingIndex >= 0) {
    state.records[existingIndex] = { ...record, id: state.records[existingIndex].id };
  } else {
    state.records.unshift(record);
  }
}

function recalcFlags() {
  const quarterRange = getQuarterRange();
  const records = state.records.map(sanitizeRecord);
  for (const record of records) {
    record.issues = [];
    if (!record.underwriterName) record.issues.push({ type: 'warn', code: 'missing_name', label: 'Missing underwriter' });
    if (!record.startDate || !record.endDate) record.issues.push({ type: 'warn', code: 'missing_dates', label: 'Missing date(s)' });
    if (record.startDate && record.endDate && record.startDate > record.endDate) {
      record.issues.push({ type: 'bad', code: 'bad_range', label: 'Start after end' });
    }
    if (activeInQuarter(record, quarterRange.start, quarterRange.end)) {
      record.issues.push({ type: 'ok', code: 'in_quarter', label: 'In quarter' });
    }
  }

  for (let i = 0; i < records.length; i += 1) {
    for (let j = i + 1; j < records.length; j += 1) {
      const a = records[i];
      const b = records[j];
      const sameHash = a.sourceHash && b.sourceHash && a.sourceHash === b.sourceHash;
      const sameUnderwriter = normalizeName(a.underwriterName) && normalizeName(a.underwriterName) === normalizeName(b.underwriterName);
      const sameRange = a.startDate && b.startDate && a.endDate && b.endDate && a.startDate === b.startDate && a.endDate === b.endDate;
      const sameAmount = String(a.amount || '') && String(a.amount) === String(b.amount);

      if (sameHash || (sameUnderwriter && sameRange && sameAmount)) {
        a.issues.push({ type: 'bad', code: 'duplicate', label: `Duplicate candidate: ${b.underwriterName || b.sourceFileName}` });
        b.issues.push({ type: 'bad', code: 'duplicate', label: `Duplicate candidate: ${a.underwriterName || a.sourceFileName}` });
      } else if (sameUnderwriter && dateRangesOverlap(a.startDate, a.endDate, b.startDate, b.endDate)) {
        a.issues.push({ type: 'warn', code: 'overlap', label: `Overlap with same underwriter: ${b.sourceFileName || b.underwriterName}` });
        b.issues.push({ type: 'warn', code: 'overlap', label: `Overlap with same underwriter: ${a.sourceFileName || a.underwriterName}` });
      }
    }
  }

  state.records = records;
}

function normalizeName(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\b(inc|llc|co|corp|corporation|company|the)\b/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function dateRangesOverlap(startA, endA, startB, endB) {
  if (!startA || !endA || !startB || !endB) return false;
  return startA <= endB && startB <= endA;
}

function activeInQuarter(record, quarterStart, quarterEnd) {
  if (!record.startDate || !record.endDate) return false;
  return record.startDate <= quarterEnd && record.endDate >= quarterStart;
}

function renderAll() {
  renderTable();
  renderSelectedRecord();
  renderMetrics();
  renderNarrative();
}

function getFilteredRecords() {
  const query = els.searchInput.value.trim().toLowerCase();
  const filter = els.issueFilter.value;
  const quarter = getQuarterRange();

  return state.records.filter((record) => {
    const haystack = [
      record.underwriterName,
      record.contractType,
      record.programName,
      record.email,
      record.phone,
      record.notes,
      record.sourceFileName,
      record.creditCopy,
      record.rawText,
    ].join(' ').toLowerCase();

    if (query && !haystack.includes(query)) return false;

    if (filter === 'issue' && !record.issues.some((issue) => issue.type !== 'ok')) return false;
    if (filter === 'duplicate' && !record.issues.some((issue) => issue.code === 'duplicate')) return false;
    if (filter === 'overlap' && !record.issues.some((issue) => issue.code === 'overlap')) return false;
    if (filter === 'quarter' && !activeInQuarter(record, quarter.start, quarter.end)) return false;
    return true;
  });
}

function renderTable() {
  const rows = getFilteredRecords();
  els.recordsBody.innerHTML = '';

  if (!rows.length) {
    els.recordsBody.innerHTML = `<tr><td colspan="8" class="muted">No matching records.</td></tr>`;
    return;
  }

  for (const record of rows) {
    const fragment = els.rowTemplate.content.cloneNode(true);
    const tr = fragment.querySelector('tr');
    if (record.id === state.selectedId) tr.classList.add('selected-row');

    fragment.querySelector('.row-underwriter').textContent = record.underwriterName || '—';
    fragment.querySelector('.row-type').textContent = record.contractType || '—';
    fragment.querySelector('.row-start').textContent = record.startDate || '—';
    fragment.querySelector('.row-end').textContent = record.endDate || '—';
    fragment.querySelector('.row-amount').textContent = formatMoney(record.amount);
    fragment.querySelector('.row-source').textContent = record.sourceFileName || '—';

    const flagCell = fragment.querySelector('.row-flags');
    flagCell.appendChild(buildFlagBadges(record.issues));

    fragment.querySelector('.row-open').addEventListener('click', () => {
      state.selectedId = record.id;
      persist();
      renderAll();
    });

    els.recordsBody.appendChild(fragment);
  }
}

function buildFlagBadges(issues) {
  const wrap = document.createElement('div');
  wrap.className = 'flag-badges';
  if (!issues.length) {
    wrap.appendChild(makeBadge('Clean', 'ok'));
    return wrap;
  }
  issues.forEach((issue) => {
    if (issue.code === 'in_quarter' && issues.length > 1) return;
    wrap.appendChild(makeBadge(issue.label, issue.type));
  });
  return wrap;
}

function makeBadge(text, type = 'ok') {
  const span = document.createElement('span');
  span.className = `badge ${type === 'bad' ? 'badge-bad' : type === 'warn' ? 'badge-warn' : 'badge-ok'}`;
  span.textContent = text;
  return span;
}

function renderSelectedRecord() {
  const record = state.records.find((item) => item.id === state.selectedId);
  if (!record) {
    els.detailBadge.textContent = 'Nothing selected';
    els.recordIssues.textContent = 'No record selected.';
    els.recordForm.reset();
    return;
  }

  els.detailBadge.textContent = record.sourceFileName || 'Manual record';
  for (const element of els.recordForm.elements) {
    if (!element.name) continue;
    element.value = record[element.name] ?? '';
  }
  els.recordIssues.textContent = record.issues.length
    ? record.issues.map((issue) => `• ${issue.label}`).join('\n')
    : 'No issues flagged for this record.';
}

function saveFormToSelected() {
  const record = state.records.find((item) => item.id === state.selectedId);
  if (!record) return;
  for (const element of els.recordForm.elements) {
    if (!element.name) continue;
    record[element.name] = element.value;
  }
  recalcFlags();
  persist();
  renderAll();
  setStatus(`Saved ${record.underwriterName || record.sourceFileName || 'record'}.`);
}

function deleteSelected() {
  const record = state.records.find((item) => item.id === state.selectedId);
  if (!record) return;
  const ok = confirm(`Delete ${record.underwriterName || record.sourceFileName || 'this record'}?`);
  if (!ok) return;
  state.records = state.records.filter((item) => item.id !== state.selectedId);
  state.selectedId = state.records[0]?.id || null;
  recalcFlags();
  persist();
  renderAll();
}

function duplicateSelected() {
  const record = state.records.find((item) => item.id === state.selectedId);
  if (!record) return;
  const clone = sanitizeRecord({
    ...structuredClone(record),
    id: crypto.randomUUID(),
    sourceHash: '',
    sourceFileName: `${record.sourceFileName || 'Manual entry'} (copy)`,
  });
  state.records.unshift(clone);
  state.selectedId = clone.id;
  recalcFlags();
  persist();
  renderAll();
}

function clearAllRecords() {
  const ok = confirm('Clear the local database? This wipes records stored in this browser.');
  if (!ok) return;
  state.records = [];
  state.selectedId = null;
  persist();
  renderAll();
  setStatus('Local database cleared.');
}

function renderMetrics() {
  const quarter = getQuarterRange();
  const inQuarter = state.records.filter((record) => activeInQuarter(record, quarter.start, quarter.end));
  const totalValue = inQuarter.reduce((sum, record) => sum + Number(record.amount || 0), 0);
  const flagged = state.records.filter((record) => record.issues.some((issue) => issue.type !== 'ok')).length;
  const duplicates = state.records.filter((record) => record.issues.some((issue) => issue.code === 'duplicate')).length;
  const uniqueUnderwriters = new Set(inQuarter.map((record) => normalizeName(record.underwriterName)).filter(Boolean)).size;

  const metrics = [
    ['Records', state.records.length],
    ['In quarter', inQuarter.length],
    ['Unique underwriters', uniqueUnderwriters],
    ['Quarter value', formatMoney(totalValue)],
    ['Flagged records', flagged],
    ['Duplicate candidates', duplicates],
  ];

  els.metrics.innerHTML = '';
  metrics.forEach(([label, value]) => {
    const card = document.createElement('div');
    card.className = 'metric';
    card.innerHTML = `<span class="metric-label">${escapeHtml(label)}</span><span class="metric-value">${escapeHtml(String(value))}</span>`;
    els.metrics.appendChild(card);
  });
}

function renderNarrative() {
  const quarter = getQuarterRange();
  const inQuarter = state.records.filter((record) => activeInQuarter(record, quarter.start, quarter.end));
  const uniqueUnderwriters = new Set(inQuarter.map((record) => normalizeName(record.underwriterName)).filter(Boolean)).size;
  const totalValue = inQuarter.reduce((sum, record) => sum + Number(record.amount || 0), 0);
  const overlapCount = state.records.filter((record) => record.issues.some((issue) => issue.code === 'overlap')).length;
  const duplicateCount = state.records.filter((record) => record.issues.some((issue) => issue.code === 'duplicate')).length;

  const lines = [
    `Quarter reviewed: ${quarter.start} through ${quarter.end}`,
    `Active contracts in quarter: ${inQuarter.length}`,
    `Unique underwriters in quarter: ${uniqueUnderwriters}`,
    `Total contract value represented in active quarter records: ${formatMoney(totalValue)}`,
    `Duplicate candidates needing review: ${duplicateCount}`,
    `Same-underwriter overlap flags needing review: ${overlapCount}`,
    '',
    'Underwriters active in quarter:',
    ...inQuarter
      .slice()
      .sort((a, b) => (a.underwriterName || '').localeCompare(b.underwriterName || ''))
      .map((record) => {
        const detail = [record.contractType, record.programName].filter(Boolean).join(' — ');
        return `- ${record.underwriterName || 'Unnamed record'} | ${record.startDate || '??'} to ${record.endDate || '??'} | ${formatMoney(record.amount)}${detail ? ` | ${detail}` : ''}`;
      }),
  ];

  els.narrativeBox.value = lines.join('\n');
}

function exportQuarterCsv() {
  const quarter = getQuarterRange();
  const rows = state.records.filter((record) => activeInQuarter(record, quarter.start, quarter.end));
  const columns = [
    'underwriterName', 'contractType', 'programName', 'contactPerson', 'email', 'phone',
    'startDate', 'endDate', 'amount', 'creditCount', 'programCount', 'sourceFileName', 'notes'
  ];
  const csv = [
    columns.join(','),
    ...rows.map((record) => columns.map((key) => csvEscape(record[key] ?? '')).join(',')),
  ].join('\n');
  downloadFile(`wnmu-underwriters-${quarter.start}-to-${quarter.end}.csv`, csv, 'text/csv;charset=utf-8');
}

function exportJsonBackup() {
  const payload = JSON.stringify({ exportedAt: new Date().toISOString(), records: state.records }, null, 2);
  downloadFile('wnmu-underwriter-backup.json', payload, 'application/json;charset=utf-8');
}

async function importJsonBackup(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  await importJsonFile(file);
  els.importJsonInput.value = '';
  recalcFlags();
  renderAll();
  persist();
}

async function importJsonFile(file) {
  const text = await file.text();
  const parsed = JSON.parse(text);
  const records = Array.isArray(parsed) ? parsed : parsed.records;
  if (!Array.isArray(records)) throw new Error('JSON backup must contain a records array.');
  records.forEach((record) => upsertRecord(sanitizeRecord(record)));
  state.selectedId = state.records[0]?.id || state.selectedId;
  setStatus(`Imported ${records.length} record(s) from JSON backup.`);
}

function copyNarrative() {
  navigator.clipboard.writeText(els.narrativeBox.value)
    .then(() => setStatus('Quarter narrative copied to clipboard.'))
    .catch(() => setStatus('Clipboard copy failed. The text is still sitting there in the box.'));
}

function getQuarterRange() {
  return {
    start: els.quarterStart.value || '2026-01-01',
    end: els.quarterEnd.value || '2026-03-31',
  };
}

function formatMoney(value) {
  const amount = Number(value || 0);
  if (!Number.isFinite(amount) || amount === 0) return '—';
  return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(amount);
}

function csvEscape(value) {
  const text = String(value ?? '');
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function setStatus(message) {
  els.importStatus.textContent = message;
}

async function hashBytes(bytes) {
  const hashBuffer = await crypto.subtle.digest('SHA-256', bytes);
  return [...new Uint8Array(hashBuffer)].map((x) => x.toString(16).padStart(2, '0')).join('');
}

function downloadFile(name, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
