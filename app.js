const CURRENT_STORAGE_KEY = 'wnmu-underwriter-intake-v0.2.0';
const LEGACY_STORAGE_KEYS = ['wnmu-underwriter-intake-v0.1.0'];
const OCR_PAGE_LIMIT = 4;

const state = {
  records: [],
  selectedId: null,
  currentTab: 'ingest',
  sortKey: 'underwriterName',
  sortDir: 'asc',
  importQueueActive: false,
};

const els = {
  tabButtons: [...document.querySelectorAll('.tab-button')],
  tabPanels: [...document.querySelectorAll('.tab-panel')],
  dropZone: document.getElementById('dropZone'),
  fileInput: document.getElementById('fileInput'),
  importStatus: document.getElementById('importStatus'),
  searchInput: document.getElementById('searchInput'),
  issueFilter: document.getElementById('issueFilter'),
  quarterStart: document.getElementById('quarterStart'),
  quarterEnd: document.getElementById('quarterEnd'),
  metrics: document.getElementById('metrics'),
  narrativeBox: document.getElementById('narrativeBox'),
  contractsBody: document.getElementById('contractsBody'),
  quarterlyBody: document.getElementById('quarterlyBody'),
  quarterSummaryBadge: document.getElementById('quarterSummaryBadge'),
  contractRowTemplate: document.getElementById('contractRowTemplate'),
  recordModal: document.getElementById('recordModal'),
  recordForm: document.getElementById('recordForm'),
  modalSubhead: document.getElementById('modalSubhead'),
  closeModalBtn: document.getElementById('closeModalBtn'),
  importJsonInput: document.getElementById('importJsonInput'),
  sortButtons: [...document.querySelectorAll('.sort-button')],
};

boot();

function boot() {
  loadState();
  bindEvents();
  recalcFlags();
  renderAll();
}

function bindEvents() {
  els.tabButtons.forEach((button) => {
    button.addEventListener('click', () => {
      state.currentTab = button.dataset.tab;
      renderTabs();
    });
  });

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
    openModal(record.id);
  });

  document.getElementById('exportJsonBtn').addEventListener('click', exportJsonBackup);
  document.getElementById('importJsonBtn').addEventListener('click', () => els.importJsonInput.click());
  document.getElementById('clearAllBtn').addEventListener('click', clearAllRecords);
  document.getElementById('exportQuarterCsvBtn').addEventListener('click', exportQuarterCsv);
  document.getElementById('copyNarrativeBtn').addEventListener('click', copyNarrative);
  document.getElementById('saveRecordBtn').addEventListener('click', (event) => {
    event.preventDefault();
    saveFormToSelected();
  });
  document.getElementById('duplicateRecordBtn').addEventListener('click', (event) => {
    event.preventDefault();
    duplicateSelected();
  });
  document.getElementById('deleteRecordBtn').addEventListener('click', (event) => {
    event.preventDefault();
    deleteSelected();
  });

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

  els.sortButtons.forEach((button) => {
    button.addEventListener('click', () => toggleSort(button.dataset.sort));
  });

  els.closeModalBtn.addEventListener('click', closeModal);
  els.recordModal.addEventListener('click', (event) => {
    if (event.target.dataset.closeModal === 'true') closeModal();
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !els.recordModal.classList.contains('hidden')) closeModal();
  });
}

function loadState() {
  let loaded = false;
  const keys = [CURRENT_STORAGE_KEY, ...LEGACY_STORAGE_KEYS];
  for (const key of keys) {
    const raw = localStorage.getItem(key);
    if (!raw) continue;
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed.records)) {
        state.records = parsed.records.map(sanitizeRecord);
        state.selectedId = parsed.selectedId || state.records[0]?.id || null;
        loaded = true;
        break;
      }
    } catch (error) {
      console.error('Failed to load local state', error);
    }
  }
  if (loaded) persist();
}

function persist() {
  localStorage.setItem(CURRENT_STORAGE_KEY, JSON.stringify({
    records: state.records,
    selectedId: state.selectedId,
  }));
}

function makeTrulyBlank() {
  return {
    id: crypto.randomUUID(),
    underwriterName: '',
    contractType: '',
    programName: '',
    placementDetail: '',
    contactPerson: '',
    email: '',
    phone: '',
    startDate: '',
    endDate: '',
    amount: '',
    creditCount: '',
    programCount: '',
    creditCopy: '',
    creditRuns: '',
    notes: '',
    rawText: '',
    sourceFileName: '',
    sourceHash: '',
    importedAt: new Date().toISOString(),
    issues: [],
    issueSummary: '',
  };
}

function makeBlankRecord() {
  return sanitizeRecord({
    ...makeTrulyBlank(),
    sourceFileName: 'Manual entry',
  });
}

function sanitizeRecord(record) {
  const blank = makeTrulyBlank();
  const merged = { ...blank, ...record, issues: Array.isArray(record?.issues) ? record.issues : [] };
  merged.issueSummary = buildIssueSummary(merged.issues);
  return merged;
}

async function importFiles(files) {
  if (!files.length) return;
  if (state.importQueueActive) {
    setStatus('Import already running. Let that finish first.');
    return;
  }
  state.importQueueActive = true;

  try {
    const accepted = files.filter((file) => /\.(pdf|docx|json)$/i.test(file.name));
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
      const record = await parseFileToRecord(file);
      upsertRecord(record);
      state.selectedId = record.id;
      recalcFlags();
      renderAll();
      persist();
    }

    state.currentTab = 'contracts';
    renderTabs();
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

  if (/\.docx$/i.test(file.name)) {
    rawText = await extractDocxText(bytes.buffer);
    extractionPath = 'DOCX text extraction';
  } else if (/\.pdf$/i.test(file.name)) {
    const pdfResult = await extractPdfText(bytes);
    rawText = pdfResult.text;
    extractionPath = pdfResult.path;
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
  const textParts = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map((item) => item.str).join(' ').trim();
    if (pageText.length > 80) textParts.push(pageText);
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
    const result = await Tesseract.recognize(canvas, 'eng', { logger: () => {} });
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
    /(?:program|series|campaign|show|special|title of program|contract title)\s*[:\-]\s*(.+)/i,
  ]) || guessProgramName(lines, lower);

  record.placementDetail = findPlacementDetail(text, lines) || '';
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
    /(?:credit copy|copy|announcement|audio credit|underwriting copy)\s*[:\-]\s*([\s\S]{20,500})/i,
  ], true) || extractQuotedBlock(text);

  if (!record.underwriterName && record.email) record.underwriterName = emailDomainToName(record.email);
  return record;
}

function detectContractType(lower, lines, fileName) {
  if (lower.includes('american revolution')) return 'American Revolution';
  if (lower.includes('underwriter') && lower.includes('contract')) return 'Underwriter Contract';
  if (lower.includes('underwriting')) return 'Underwriting';
  const line = lines.find((item) => /contract|agreement|underwriting|sponsorship/i.test(item));
  if (line) return line.slice(0, 80);
  return fileName.replace(/\.[^.]+$/, '');
}

function findPlacementDetail(text, lines) {
  const direct = findNamedField(text, [
    /(?:day[ -]?part|daypart)\s*[:\-]\s*(.+)/i,
    /(?:placement|schedule|program\/day\/day-part|program\s*\/\s*day\s*\/\s*day-part)\s*[:\-]\s*(.+)/i,
    /(?:run(?:s)? on|air(?:s|ing)? on)\s*[:\-]\s*(.+)/i,
    /(?:program and daypart|program daypart)\s*[:\-]\s*(.+)/i,
  ]);
  if (direct) return direct;

  const candidate = lines.find((line) => {
    const lower = line.toLowerCase();
    return line.length < 120 && /(monday|tuesday|wednesday|thursday|friday|saturday|sunday|weekday|weekend|morning|afternoon|evening|overnight|drive|news hour|edition|program)/i.test(lower);
  });
  return candidate || '';
}

function guessProgramName(lines, lower) {
  const candidates = lines.filter((line) => {
    const l = line.toLowerCase();
    return l.length > 5 && l.length < 90 && !/(wnmu|contract|agreement|phone|email|address|date|amount|signature)/i.test(l);
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

function extractQuotedBlock(text) {
  const match = text.match(/["“]([^"”]{20,500})["”]/);
  return match?.[1]?.trim() || '';
}

function emailDomainToName(email) {
  const domain = (email.split('@')[1] || '').replace(/\.[a-z]{2,}$/i, '');
  return domain
    .split(/[.-]/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(' ');
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
    state.records[existingIndex] = sanitizeRecord({ ...record, id: state.records[existingIndex].id });
  } else {
    state.records.unshift(sanitizeRecord(record));
  }
}

function recalcFlags() {
  const quarterRange = getQuarterRange();
  const records = state.records.map(sanitizeRecord);

  for (const record of records) {
    record.issues = [];
    if (!record.underwriterName) record.issues.push({ type: 'warn', code: 'missing_name', label: 'Missing name' });
    if (!record.startDate || !record.endDate) record.issues.push({ type: 'warn', code: 'missing_dates', label: 'Missing date(s)' });
    if (record.startDate && record.endDate && record.startDate > record.endDate) {
      record.issues.push({ type: 'bad', code: 'bad_range', label: 'Start after end' });
    }
    if (!record.placementDetail) record.issues.push({ type: 'warn', code: 'missing_placement', label: 'Missing placement' });
    if (!record.creditRuns) record.issues.push({ type: 'warn', code: 'missing_runs', label: 'Missing exact runs' });
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
        a.issues.push({ type: 'warn', code: 'overlap', label: `Overlap with same name: ${b.sourceFileName || b.underwriterName}` });
        b.issues.push({ type: 'warn', code: 'overlap', label: `Overlap with same name: ${a.sourceFileName || a.underwriterName}` });
      }
    }
  }

  for (const record of records) record.issueSummary = buildIssueSummary(record.issues);
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

function getFilteredRecords() {
  const query = els.searchInput.value.trim().toLowerCase();
  const filter = els.issueFilter.value;
  const quarter = getQuarterRange();

  return state.records.filter((record) => {
    const haystack = [
      record.underwriterName,
      record.contractType,
      record.programName,
      record.placementDetail,
      record.email,
      record.phone,
      record.notes,
      record.sourceFileName,
      record.creditCopy,
      record.creditRuns,
      record.rawText,
    ].join(' ').toLowerCase();

    if (query && !haystack.includes(query)) return false;
    if (filter === 'issue' && !record.issues.some((issue) => issue.type !== 'ok')) return false;
    if (filter === 'duplicate' && !record.issues.some((issue) => issue.code === 'duplicate')) return false;
    if (filter === 'overlap' && !record.issues.some((issue) => issue.code === 'overlap')) return false;
    if (filter === 'quarter' && !activeInQuarter(record, quarter.start, quarter.end)) return false;
    if (filter === 'missing_runs' && record.creditRuns) return false;
    return true;
  });
}

function getSortedRecords(records) {
  const list = [...records];
  const dir = state.sortDir === 'asc' ? 1 : -1;
  list.sort((a, b) => compareForSort(a, b, state.sortKey) * dir);
  return list;
}

function compareForSort(a, b, key) {
  switch (key) {
    case 'amount':
      return Number(a.amount || 0) - Number(b.amount || 0);
    case 'startDate':
    case 'endDate':
      return String(a[key] || '9999-99-99').localeCompare(String(b[key] || '9999-99-99'));
    case 'flags':
      return flagScore(b) - flagScore(a);
    default:
      return String(a[key] || '').localeCompare(String(b[key] || ''), undefined, { sensitivity: 'base' });
  }
}

function flagScore(record) {
  return record.issues.reduce((score, issue) => {
    if (issue.type === 'bad') return score + 100;
    if (issue.type === 'warn') return score + 10;
    return score + 1;
  }, 0);
}

function toggleSort(key) {
  if (state.sortKey === key) {
    state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
  } else {
    state.sortKey = key;
    state.sortDir = key === 'amount' ? 'desc' : 'asc';
  }
  renderContractsTable();
  renderSortButtons();
}

function renderAll() {
  renderTabs();
  renderSortButtons();
  renderContractsTable();
  renderQuarterly();
  renderMetrics();
  renderNarrative();
  syncModal();
}

function renderTabs() {
  els.tabButtons.forEach((button) => button.classList.toggle('active', button.dataset.tab === state.currentTab));
  els.tabPanels.forEach((panel) => panel.classList.toggle('active', panel.id === `tab-${state.currentTab}`));
}

function renderSortButtons() {
  els.sortButtons.forEach((button) => {
    const isActive = button.dataset.sort === state.sortKey;
    button.classList.toggle('active', isActive);
    const arrow = button.querySelector('.sort-arrow');
    if (arrow) arrow.textContent = isActive ? (state.sortDir === 'asc' ? '▲' : '▼') : '↕';
  });
}

function renderContractsTable() {
  const rows = getSortedRecords(getFilteredRecords());
  els.contractsBody.innerHTML = '';

  if (!rows.length) {
    els.contractsBody.innerHTML = '<tr><td colspan="9" class="muted">No matching records.</td></tr>';
    return;
  }

  for (const record of rows) {
    const fragment = els.contractRowTemplate.content.cloneNode(true);
    fragment.querySelector('.row-underwriter').textContent = record.underwriterName || '—';
    fragment.querySelector('.row-start').textContent = record.startDate || '—';
    fragment.querySelector('.row-end').textContent = record.endDate || '—';
    fragment.querySelector('.row-type').textContent = record.contractType || '—';
    fragment.querySelector('.row-placement').textContent = record.placementDetail || '—';
    fragment.querySelector('.row-amount').textContent = formatMoney(record.amount);
    fragment.querySelector('.row-source').textContent = record.sourceFileName || '—';
    fragment.querySelector('.row-flags').appendChild(buildFlagBadges(record.issues));
    fragment.querySelector('.row-open').addEventListener('click', () => openModal(record.id));
    els.contractsBody.appendChild(fragment);
  }
}

function renderQuarterly() {
  const quarter = getQuarterRange();
  const rows = getSortedRecords(state.records.filter((record) => activeInQuarter(record, quarter.start, quarter.end)));
  els.quarterlyBody.innerHTML = '';
  els.quarterSummaryBadge.textContent = `${rows.length} record${rows.length === 1 ? '' : 's'}`;

  if (!rows.length) {
    els.quarterlyBody.innerHTML = '<tr><td colspan="6" class="muted">No records active in the selected quarter.</td></tr>';
    return;
  }

  for (const record of rows) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escapeHtml(record.underwriterName || '—')}</td>
      <td>${escapeHtml(record.placementDetail || record.programName || '—')}</td>
      <td>${escapeHtml(record.startDate || '—')}</td>
      <td>${escapeHtml(record.endDate || '—')}</td>
      <td>${escapeHtml(formatMoney(record.amount))}</td>
      <td>${escapeHtml(record.creditRuns || '—')}</td>
    `;
    els.quarterlyBody.appendChild(tr);
  }
}

function renderMetrics() {
  const quarter = getQuarterRange();
  const inQuarter = state.records.filter((record) => activeInQuarter(record, quarter.start, quarter.end));
  const totalValue = inQuarter.reduce((sum, record) => sum + Number(record.amount || 0), 0);
  const flagged = state.records.filter((record) => record.issues.some((issue) => issue.type !== 'ok')).length;
  const duplicates = state.records.filter((record) => record.issues.some((issue) => issue.code === 'duplicate')).length;
  const uniqueUnderwriters = new Set(inQuarter.map((record) => normalizeName(record.underwriterName)).filter(Boolean)).size;
  const missingRuns = inQuarter.filter((record) => !record.creditRuns).length;
  const metrics = [
    ['Records', state.records.length],
    ['In quarter', inQuarter.length],
    ['Unique names', uniqueUnderwriters],
    ['Quarter value', formatMoney(totalValue)],
    ['Flagged records', flagged],
    ['Missing exact runs', missingRuns],
    ['Duplicate candidates', duplicates],
    ['Sort mode', `${labelForSort(state.sortKey)} ${state.sortDir}`],
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
  const inQuarter = getSortedRecords(state.records.filter((record) => activeInQuarter(record, quarter.start, quarter.end)));
  const uniqueUnderwriters = new Set(inQuarter.map((record) => normalizeName(record.underwriterName)).filter(Boolean)).size;
  const totalValue = inQuarter.reduce((sum, record) => sum + Number(record.amount || 0), 0);
  const overlapCount = state.records.filter((record) => record.issues.some((issue) => issue.code === 'overlap')).length;
  const duplicateCount = state.records.filter((record) => record.issues.some((issue) => issue.code === 'duplicate')).length;
  const missingRuns = inQuarter.filter((record) => !record.creditRuns).length;

  const lines = [
    `Quarter reviewed: ${quarter.start} through ${quarter.end}`,
    `Active contracts in quarter: ${inQuarter.length}`,
    `Unique underwriters in quarter: ${uniqueUnderwriters}`,
    `Total contract value represented in active quarter records: ${formatMoney(totalValue)}`,
    `Duplicate candidates needing review: ${duplicateCount}`,
    `Same-underwriter overlap flags needing review: ${overlapCount}`,
    `Quarter records still missing exact credit run dates / times: ${missingRuns}`,
    '',
    'Quarter detail:',
    ...inQuarter.map((record) => {
      const parts = [
        record.underwriterName || 'Unnamed record',
        `${record.startDate || '??'} to ${record.endDate || '??'}`,
        formatMoney(record.amount),
        record.placementDetail || record.programName || 'Placement blank',
        record.creditRuns ? `Runs entered` : `Runs still blank`,
      ];
      return `- ${parts.join(' | ')}`;
    }),
  ];

  els.narrativeBox.value = lines.join('\n');
}

function openModal(recordId) {
  state.selectedId = recordId;
  syncModal();
  els.recordModal.classList.remove('hidden');
  els.recordModal.setAttribute('aria-hidden', 'false');
}

function closeModal() {
  els.recordModal.classList.add('hidden');
  els.recordModal.setAttribute('aria-hidden', 'true');
}

function syncModal() {
  const record = state.records.find((item) => item.id === state.selectedId);
  if (!record) {
    els.modalSubhead.textContent = 'No record selected';
    els.recordForm.reset();
    return;
  }

  els.modalSubhead.textContent = record.sourceFileName || 'Manual record';
  for (const element of els.recordForm.elements) {
    if (!element.name) continue;
    if (element.name === 'issueSummary') {
      element.value = record.issueSummary || 'No flags';
      continue;
    }
    element.value = record[element.name] ?? '';
  }
}

function saveFormToSelected() {
  const record = state.records.find((item) => item.id === state.selectedId);
  if (!record) return;
  for (const element of els.recordForm.elements) {
    if (!element.name || element.name === 'issueSummary') continue;
    record[element.name] = element.value;
  }
  recalcFlags();
  persist();
  renderAll();
  syncModal();
  setStatus(`Saved ${record.underwriterName || record.sourceFileName || 'record'}.`);
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
  syncModal();
  setStatus('Record duplicated.');
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
  syncModal();
  if (!state.selectedId) closeModal();
}

function clearAllRecords() {
  const ok = confirm('Clear the local database? This wipes records stored in this browser.');
  if (!ok) return;
  state.records = [];
  state.selectedId = null;
  persist();
  renderAll();
  closeModal();
  setStatus('Local database cleared.');
}

function exportQuarterCsv() {
  const quarter = getQuarterRange();
  const rows = getSortedRecords(state.records.filter((record) => activeInQuarter(record, quarter.start, quarter.end)));
  const columns = [
    'underwriterName',
    'startDate',
    'endDate',
    'issueSummary',
    'contractType',
    'placementDetail',
    'amount',
    'sourceFileName',
    'creditRuns',
    'notes',
  ];
  const csv = [
    columns.join(','),
    ...rows.map((record) => columns.map((key) => csvEscape(record[key] ?? '')).join(',')),
  ].join('\n');
  downloadFile(`wnmu-quarterly-underwriters-${quarter.start}-to-${quarter.end}.csv`, csv, 'text/csv;charset=utf-8');
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

function buildFlagBadges(issues) {
  const wrap = document.createElement('div');
  wrap.className = 'flag-badges';
  const visible = issues.filter((issue) => !(issue.code === 'in_quarter' && issues.length > 1));
  if (!visible.length) {
    wrap.appendChild(makeBadge('Clean', 'ok'));
    return wrap;
  }
  visible.forEach((issue) => wrap.appendChild(makeBadge(issue.label, issue.type)));
  return wrap;
}

function makeBadge(text, type = 'ok') {
  const span = document.createElement('span');
  span.className = `badge ${type === 'bad' ? 'badge-bad' : type === 'warn' ? 'badge-warn' : 'badge-ok'}`;
  span.textContent = text;
  return span;
}

function buildIssueSummary(issues) {
  const visible = (issues || []).filter((issue) => issue.code !== 'in_quarter');
  return visible.length ? visible.map((issue) => issue.label).join(' | ') : 'No flags';
}

function labelForSort(key) {
  const labels = {
    underwriterName: 'underwriter',
    startDate: 'start',
    endDate: 'end',
    flags: 'flags',
    contractType: 'type',
    placementDetail: 'placement',
    amount: 'amount',
    sourceFileName: 'source',
  };
  return labels[key] || key;
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
