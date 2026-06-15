/* ─── State ─────────────────────────────────────────── */
let rawData = [];
let currentStage = 'ثانوية';
let currentClass = 'all';
let currentExam  = 'all';
let extraClasses = { 'ثانوية': [], 'متوسطة': [] };

let barChartInst  = null;
let pieChartInst  = null;
let lineChartInst = null;

const EXAM_COLS   = ['فتري1', 'فتري2', 'نهائي'];
const EXAM_LABELS = { 'فتري1': 'فتري 1', 'فتري2': 'فتري 2', 'نهائي': 'نهائي' };

const GRADES = [
  { label: 'ممتاز',    min: 90, color: '#10b981' },
  { label: 'جيد جداً', min: 80, color: '#3b82f6' },
  { label: 'جيد',      min: 70, color: '#f59e0b' },
  { label: 'مقبول',   min: 60, color: '#f97316' },
  { label: 'راسب',     min: 0,  color: '#ef4444' },
];

const STAGE_COLORS = { 'ثانوية': '#1e3a5f', 'متوسطة': '#1a5e3a' };

/* ─── Drag & Drop ───────────────────────────────────── */
function handleDragOver(e)  { e.preventDefault(); document.getElementById('dropZone').classList.add('drop-zone-active'); }
function handleDragLeave(e) { document.getElementById('dropZone').classList.remove('drop-zone-active'); }
function handleDrop(e)      { e.preventDefault(); document.getElementById('dropZone').classList.remove('drop-zone-active'); const f = e.dataTransfer.files[0]; if (f) processFile(f); }
function handleFileInput(e) { const f = e.target.files[0]; if (f) processFile(f); }

/* ─── File Processing ───────────────────────────────── */
function processFile(file) {
  const name = file.name.toLowerCase();
  const el = document.getElementById('fileNameDisplay');
  el.textContent = '📄 ' + file.name;
  el.classList.remove('hidden');

  if (name.endsWith('.csv')) {
    Papa.parse(file, { header: true, skipEmptyLines: true,
      complete: r => initData(r.data),
      error: e => alert('خطأ: ' + e.message) });
  } else if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      initData(XLSX.utils.sheet_to_json(ws, { defval: '' }));
    };
    reader.readAsArrayBuffer(file);
  } else if (name.endsWith('.pdf')) {
    processPDF(file);
  } else {
    alert('صيغة غير مدعومة. استخدم CSV أو Excel أو PDF.');
  }
}

/* ─── PDF Processing ────────────────────────────────── */
let pdfExtractedRows = [];
let pdfHeaders       = [];

// Convert Eastern Arabic-Indic digits → Western digits
function normalizeNumerals(s) {
  return String(s).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d));
}

function setProgress(pct, msg) {
  document.getElementById('pdfProgressBar').style.width = pct + '%';
  document.getElementById('pdfProgressText').textContent = msg;
}

async function processPDF(file) {
  const progress = document.getElementById('pdfProgress');
  progress.classList.remove('hidden');
  setProgress(3, 'جاري فتح الملف…');

  try {
    const pdfjsLib = window['pdfjs-dist/build/pdf'];
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';

    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    // ── Step 1: Try native text extraction ──────────────
    setProgress(10, 'استخراج النص من PDF…');
    const textRows = await extractTextRows(pdf);

    if (textRows.length >= 2) {
      // PDF has readable text — use it directly
      setProgress(95, 'تحليل هيكل الجدول…');
      finalizePdfRows(textRows);
      setProgress(100, 'اكتمل!');
      setTimeout(() => progress.classList.add('hidden'), 400);
      openPdfMapper();
      return;
    }

    // ── Step 2: PDF is scanned — fall back to OCR ───────
    setProgress(15, 'الملف يحتوي صوراً — جاري تحميل محرك OCR العربي…');

    const worker = await Tesseract.createWorker('ara+eng', 1, {
      logger: m => {
        if (m.status === 'loading tesseract core') setProgress(20, 'تحميل محرك OCR…');
        if (m.status === 'initializing tesseract')  setProgress(30, 'تهيئة المحرك…');
        if (m.status === 'loading language traineddata') setProgress(40, 'تحميل بيانات اللغة العربية…');
        if (m.status === 'recognizing text') {
          const pct = Math.round(50 + (m.progress ?? 0) * 40);
          setProgress(pct, `جاري التعرف على النص… ${Math.round((m.progress ?? 0) * 100)}%`);
        }
      },
    });

    // Render each page to canvas and OCR it
    let fullText = '';
    for (let p = 1; p <= pdf.numPages; p++) {
      setProgress(
        Math.round(50 + ((p - 1) / pdf.numPages) * 40),
        `تحليل صفحة ${p} من ${pdf.numPages} بالذكاء الاصطناعي…`
      );
      const canvas = await renderPageToCanvas(pdf, p, 2.5);
      const { data: { text } } = await worker.recognize(canvas);
      fullText += '\n' + text;
    }
    await worker.terminate();

    setProgress(95, 'تحليل هيكل الجدول…');
    const ocrRows = parseOcrText(fullText);

    if (!ocrRows.length) {
      progress.classList.add('hidden');
      alert('لم يتمكن النظام من استخراج بيانات جدولية من الملف.\nتأكد أن الصورة واضحة وغير مائلة.');
      return;
    }

    finalizePdfRows(ocrRows);
    setProgress(100, 'اكتمل!');
    setTimeout(() => progress.classList.add('hidden'), 400);
    openPdfMapper();

  } catch (err) {
    document.getElementById('pdfProgress').classList.add('hidden');
    console.error(err);
    alert('حدث خطأ أثناء معالجة PDF:\n' + err.message);
  }
}

/* ── Native text extraction ──────────────────────────── */
async function extractTextRows(pdf) {
  const allRows = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page    = await pdf.getPage(p);
    const content = await page.getTextContent();

    const items = content.items
      .map(i => ({ text: i.str.trim(), x: Math.round(i.transform[4]), y: Math.round(i.transform[5]) }))
      .filter(i => i.text);

    if (!items.length) continue;

    items.sort((a, b) => b.y - a.y || a.x - b.x);

    let cur = [], lastY = null;
    for (const it of items) {
      if (lastY === null || Math.abs(it.y - lastY) <= 6) {
        cur.push(it);
        if (lastY === null) lastY = it.y;
      } else {
        if (cur.length) allRows.push(cur.map(i => normalizeNumerals(i.text)).filter(Boolean));
        cur = [it]; lastY = it.y;
      }
    }
    if (cur.length) allRows.push(cur.map(i => normalizeNumerals(i.text)).filter(Boolean));
  }
  return allRows.filter(r => r.length >= 2);
}

/* ── Render PDF page to canvas for OCR ──────────────── */
async function renderPageToCanvas(pdf, pageNum, scale = 2.5) {
  const page     = await pdf.getPage(pageNum);
  const viewport = page.getViewport({ scale });
  const canvas   = document.createElement('canvas');
  canvas.width   = viewport.width;
  canvas.height  = viewport.height;
  const ctx      = canvas.getContext('2d');
  // White background for better OCR
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: ctx, viewport }).promise;
  return canvas;
}

/* ── Parse raw OCR text into rows ────────────────────── */
function parseOcrText(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const rows  = [];

  for (const line of lines) {
    // Split on 2+ spaces, tabs, or pipe chars (common OCR table separators)
    const cells = line.split(/\s{2,}|\t+|\|/).map(c => normalizeNumerals(c.trim())).filter(Boolean);
    if (cells.length >= 2) rows.push(cells);
  }

  // Merge adjacent single-word lines that form one row (OCR sometimes splits)
  return rows;
}

/* ── Detect header and data rows ─────────────────────── */
function finalizePdfRows(allRows) {
  const headerKeywords = ['اسم','طالب','فصل','درجة','اختبار','فتري','نهائي','مرحلة','شعبة','رقم','م'];
  let headerIdx = allRows.findIndex(row =>
    row.some(cell => headerKeywords.some(k => cell.includes(k)))
  );
  if (headerIdx === -1) headerIdx = 0;

  pdfHeaders       = allRows[headerIdx].map((h, i) => h || `عمود ${i + 1}`);
  pdfExtractedRows = allRows.slice(headerIdx + 1).filter(row =>
    row.some(cell => /\d/.test(cell))
  );
}

/* ─── PDF Simple Mapper ─────────────────────────────── */
let pdfSelectedStage = 'ثانوية';
let pdfSelectedExam  = 'نهائي';

function openPdfMapper() {
  const modal = document.getElementById('pdfMapperModal');
  modal.classList.remove('hidden');
  modal.classList.add('modal-open');

  document.getElementById('pdfRowCount').textContent =
    `تم استخراج ${pdfExtractedRows.length} سجل من الملف`;

  // Reset UI state
  pdfSelectedStage = 'ثانوية';
  pdfSelectedExam  = 'نهائي';
  setPdfStage('ثانوية');
  setPdfExam('نهائي');
  document.getElementById('pdfClassInput').value = '';

  // Detect numeric columns → score column selector
  const numericCols = pdfHeaders
    .map((h, i) => ({ h, i }))
    .filter(({ i }) => pdfExtractedRows.slice(0, 10).filter(r => {
      const v = parseFloat(normalizeNumerals(r[i] ?? ''));
      return !isNaN(v) && v >= 0 && v <= 150;
    }).length >= 2);

  const wrap = document.getElementById('pdfScoreColWrap');
  const sel  = document.getElementById('pdfScoreCol');

  if (numericCols.length > 1) {
    sel.innerHTML = numericCols.map(({ h, i }) =>
      `<option value="${i}">${h || 'عمود ' + (i + 1)}</option>`
    ).join('');
    // Auto-select column with keyword درجة/نهائي/فتري
    const best = numericCols.find(({ h }) =>
      ['درجة','نهائي','فتري','score','grade'].some(k => h.includes(k))
    );
    if (best) sel.value = best.i;
    wrap.classList.remove('hidden');
  } else {
    wrap.classList.add('hidden');
    if (numericCols.length === 1) sel.value = numericCols[0].i;
    else sel.value = pdfHeaders.length - 1;
  }
}

function setPdfStage(stage) {
  pdfSelectedStage = stage;
  document.getElementById('pdfBtnSec').className =
    `pdf-stage-btn flex-1 py-2.5 rounded-xl text-sm font-bold border-2 transition-all ${stage === 'ثانوية' ? 'border-blue-600 bg-blue-600 text-white' : 'border-slate-200 bg-white text-slate-600'}`;
  document.getElementById('pdfBtnMid').className =
    `pdf-stage-btn flex-1 py-2.5 rounded-xl text-sm font-bold border-2 transition-all ${stage === 'متوسطة' ? 'border-emerald-600 bg-emerald-600 text-white' : 'border-slate-200 bg-white text-slate-600'}`;
}

function setPdfExam(exam) {
  pdfSelectedExam = exam;
  const map = { 'فتري1':'pdfExamF1', 'فتري2':'pdfExamF2', 'نهائي':'pdfExamFn' };
  Object.entries(map).forEach(([k, id]) => {
    const active = k === exam;
    document.getElementById(id).className =
      `pdf-exam-btn flex-1 py-2 rounded-xl text-sm font-medium border-2 transition-all ${active ? 'border-blue-600 bg-blue-600 text-white' : 'border-slate-200 bg-white text-slate-600'}`;
  });
}

function closePdfMapper() {
  const modal = document.getElementById('pdfMapperModal');
  modal.classList.add('hidden');
  modal.classList.remove('modal-open');
}

function applyPdfMapping() {
  const classVal = document.getElementById('pdfClassInput').value.trim();
  if (!classVal) {
    alert('يرجى إدخال رقم الصف أو الفصل.');
    document.getElementById('pdfClassInput').focus();
    return;
  }

  const scoreColIdx = parseInt(document.getElementById('pdfScoreCol').value ?? '-1');

  // Extract scores from the detected column (or scan all columns for valid scores)
  const scores = pdfExtractedRows.map((row, i) => {
    let score = NaN;
    if (scoreColIdx >= 0 && row[scoreColIdx] !== undefined) {
      score = parseFloat(normalizeNumerals(row[scoreColIdx]));
    }
    // Fallback: find first numeric value in 0-100 range in any cell
    if (isNaN(score) || score < 0 || score > 100) {
      for (const cell of row) {
        const v = parseFloat(normalizeNumerals(cell));
        if (!isNaN(v) && v >= 0 && v <= 100) { score = v; break; }
      }
    }
    // Name: first Arabic-text cell
    const name = row.find(c => /[؀-ۿ]/.test(c) && c.length > 1) ?? `طالب ${i + 1}`;
    return { name, score };
  }).filter(r => !isNaN(r.score));

  if (!scores.length) {
    alert('لم يتم العثور على درجات رقمية في الملف.');
    return;
  }

  // Build rawData entries
  const examKey = pdfSelectedExam;
  const newRows = scores.map((r, i) => ({
    id: i + 1,
    name:  r.name,
    stage: pdfSelectedStage,
    class: classVal,
    فتري1: examKey === 'فتري1' ? r.score : 0,
    فتري2: examKey === 'فتري2' ? r.score : 0,
    نهائي: examKey === 'نهائي' ? r.score : 0,
  }));

  closePdfMapper();

  // Merge with existing data (don't wipe other stages/classes)
  rawData = rawData.filter(r => !(r.stage === pdfSelectedStage && r.class === classVal));
  rawData.push(...newRows);

  // Initialize extraClasses if needed
  if (!extraClasses[pdfSelectedStage]) extraClasses[pdfSelectedStage] = [];

  // switchStage resets currentClass — set it again after
  document.getElementById('stageSection').classList.remove('hidden');
  switchStage(pdfSelectedStage);
  currentClass = classVal;
  currentExam  = examKey;
  renderClassChips();
  applyFilters();

  // Show grade distribution
  updateGradeDist(scores.map(r => r.score), `${pdfSelectedStage} — الفصل ${classVal} — ${EXAM_LABELS[examKey]}`);
}

/* ─── Column Detection ──────────────────────────────── */
function detectColumns(keys) {
  const low = s => String(s).toLowerCase().trim();
  const find = patterns => keys.find(k => patterns.some(p => low(k).includes(p)));
  return {
    name:  find(['اسم', 'name', 'طالب']),
    stage: find(['مرحلة', 'stage', 'level']),
    class: find(['فصل', 'class', 'شعبة', 'section']),
    exam1: find(['فتري1', 'فتري 1', 'mid1', 'first']),
    exam2: find(['فتري2', 'فتري 2', 'mid2', 'second']),
    final: find(['نهائي', 'final', 'نهاية']),
  };
}

/* ─── Init Data ─────────────────────────────────────── */
function initData(data) {
  if (!data || !data.length) { alert('الملف فارغ.'); return; }
  const keys = Object.keys(data[0]);
  const cols = detectColumns(keys);

  rawData = data.map((row, i) => ({
    id:    i + 1,
    name:  String(row[cols.name]  || row[keys[0]] || '—').trim(),
    stage: String(row[cols.stage] || row[keys[1]] || 'ثانوية').trim(),
    class: String(row[cols.class] || row[keys[2]] || '1').trim(),
    فتري1: parseFloat(row[cols.exam1] || row[keys[3]]) || 0,
    فتري2: parseFloat(row[cols.exam2] || row[keys[4]]) || 0,
    نهائي: parseFloat(row[cols.final] || row[keys[5]]) || 0,
  }));

  currentStage = 'ثانوية';
  currentClass = 'all';
  currentExam  = 'all';
  extraClasses = { 'ثانوية': [], 'متوسطة': [] };

  document.getElementById('stageSection').classList.remove('hidden');
  switchStage('ثانوية');
}

/* ─── Stage Switching ───────────────────────────────── */
function switchStage(stage) {
  currentStage = stage;
  currentClass = 'all';

  const isSec = stage === 'ثانوية';
  const tabSec = document.getElementById('tabSec');
  const tabMid = document.getElementById('tabMid');

  tabSec.className = `tab-active-sec flex-1 sm:flex-none px-6 py-2.5 rounded-xl font-bold text-sm transition-all border shadow-sm flex items-center justify-center gap-2 ${isSec ? '' : 'border-slate-200 bg-white text-slate-600'}`;
  tabMid.className = `flex-1 sm:flex-none px-6 py-2.5 rounded-xl font-bold text-sm transition-all border shadow-sm flex items-center justify-center gap-2 ${!isSec ? 'tab-active-mid' : 'border-slate-200 bg-white text-slate-600'}`;

  document.getElementById('stageLabel').textContent = `فصول المرحلة ${stage}`;
  renderClassChips();
  applyFilters();
}

/* ─── Class Chips ───────────────────────────────────── */
function renderClassChips() {
  const stageData   = rawData.filter(r => r.stage === currentStage);
  const baseClasses = [...new Set(stageData.map(r => r.class))].sort((a,b) => +a - +b || a.localeCompare(b, 'ar'));
  const allClasses  = [...baseClasses, ...extraClasses[currentStage]];

  const isSec  = currentStage === 'ثانوية';
  const active = isSec ? 'chip-active-sec' : 'chip-active-mid';

  const container = document.getElementById('classChips');
  container.innerHTML = '';

  // "الكل" chip — no actions
  const all = document.createElement('button');
  all.className = `class-chip px-4 py-2 rounded-xl text-sm font-semibold border ${currentClass === 'all' ? active : 'border-slate-200 bg-white text-slate-600'}`;
  all.textContent = 'جميع الفصول';
  all.onclick = () => { currentClass = 'all'; renderClassChips(); applyFilters(); };
  container.appendChild(all);

  allClasses.forEach(cls => {
    const isActive = currentClass === cls;

    const wrap = document.createElement('div');
    wrap.className = `class-chip group relative flex items-center gap-1 rounded-xl border text-sm font-semibold transition-all ${isActive ? active : 'border-slate-200 bg-white text-slate-600'}`;

    // Main label
    const label = document.createElement('button');
    label.className = 'px-3 py-2 flex-1';
    label.textContent = `الفصل ${cls}`;
    label.onclick = () => { currentClass = cls; renderClassChips(); applyFilters(); };

    // Action buttons — visible on hover or when chip is active
    const actions = document.createElement('div');
    actions.className = `flex items-center gap-0.5 pl-2 pr-1 opacity-0 group-hover:opacity-100 transition-opacity ${isActive ? '!opacity-100' : ''}`;

    const editBtn = document.createElement('button');
    editBtn.title = 'تعديل';
    editBtn.className = 'p-1 rounded-lg hover:bg-white/30 transition-colors';
    editBtn.innerHTML = `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/></svg>`;
    editBtn.onclick = (e) => { e.stopPropagation(); openEditClassModal(cls); };

    const delBtn = document.createElement('button');
    delBtn.title = 'حذف';
    delBtn.className = 'p-1 rounded-lg hover:bg-red-500/20 text-red-400 hover:text-red-600 transition-colors';
    delBtn.innerHTML = `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>`;
    delBtn.onclick = (e) => { e.stopPropagation(); confirmDeleteClass(cls); };

    actions.appendChild(editBtn);
    actions.appendChild(delBtn);
    wrap.appendChild(label);
    wrap.appendChild(actions);
    container.appendChild(wrap);
  });
}

/* ─── Delete Class ──────────────────────────────────── */
function confirmDeleteClass(cls) {
  const count = rawData.filter(r => r.stage === currentStage && r.class === cls).length;
  const msg   = count > 0
    ? `هل تريد حذف الفصل ${cls}؟\nسيتم إزالة ${count} طالب من التحليل.`
    : `هل تريد حذف الفصل ${cls}؟`;
  if (!confirm(msg)) return;
  // Remove from rawData
  rawData = rawData.filter(r => !(r.stage === currentStage && r.class === cls));
  // Remove from extraClasses if present
  extraClasses[currentStage] = extraClasses[currentStage].filter(c => c !== cls);
  if (currentClass === cls) currentClass = 'all';
  renderClassChips();
  applyFilters();
}

/* ─── Edit Class Modal ──────────────────────────────── */
let editingClass = null;

function openEditClassModal(cls) {
  editingClass = cls;
  document.getElementById('editClassName').value = cls;
  const m = document.getElementById('editClassModal');
  m.classList.remove('hidden');
  m.classList.add('modal-open');
  setTimeout(() => document.getElementById('editClassName').focus(), 50);
}

function closeEditClassModal() {
  const m = document.getElementById('editClassModal');
  m.classList.add('hidden');
  m.classList.remove('modal-open');
  editingClass = null;
}

function confirmEditClass() {
  const newName = document.getElementById('editClassName').value.trim();
  if (!newName) { alert('أدخل اسم الفصل.'); return; }
  if (newName === editingClass) { closeEditClassModal(); return; }

  const stageData = rawData.filter(r => r.stage === currentStage);
  const existing  = [...new Set(stageData.map(r => r.class)), ...extraClasses[currentStage]];
  if (existing.includes(newName)) { alert('هذا الاسم مستخدم مسبقاً.'); return; }

  // Rename in rawData
  rawData.forEach(r => { if (r.stage === currentStage && r.class === editingClass) r.class = newName; });
  // Rename in extraClasses if present
  const idx = extraClasses[currentStage].indexOf(editingClass);
  if (idx !== -1) extraClasses[currentStage][idx] = newName;
  // Keep selection
  if (currentClass === editingClass) currentClass = newName;

  closeEditClassModal();
  renderClassChips();
  applyFilters();
}

/* ─── Exam Filter Buttons ───────────────────────────── */
function setExamFilter(exam) {
  currentExam = exam;
  const map = { all:'ef-all', 'فتري1':'ef-f1', 'فتري2':'ef-f2', 'نهائي':'ef-fn' };
  document.querySelectorAll('.exam-btn').forEach(b => {
    b.className = 'exam-btn px-4 py-1.5 rounded-lg text-sm font-medium border border-slate-200 text-slate-600 bg-white hover:bg-slate-50 transition-all';
  });
  const active = document.getElementById(map[exam]);
  if (active) active.className = 'exam-btn active-exam-btn px-4 py-1.5 rounded-lg text-sm font-medium border border-blue-600 bg-blue-600 text-white transition-all';
  applyFilters();
}

/* ─── Core Filter & Render ──────────────────────────── */
function applyFilters() {
  let data = rawData.filter(r => r.stage === currentStage);
  if (currentClass !== 'all') data = data.filter(r => r.class === currentClass);

  if (!data.length) { hideSections(); return; }

  showSections();
  const scores = getScores(data, currentExam);
  updateStats(data, scores);
  updateGradeDist(scores,
    `${currentStage}${currentClass !== 'all' ? ' — الفصل ' + currentClass : ''} — ${currentExam === 'all' ? 'جميع الاختبارات' : EXAM_LABELS[currentExam]}`
  );
  updateBarChart(data, currentExam);
  updatePieChart(scores);
  updateLineChart(rawData.filter(r => r.stage === currentStage));
}

function hideSections() {
  ['statsSection','gradeDistSection','chartsSection','examFilterSection'].forEach(id =>
    document.getElementById(id).classList.add('hidden'));
}

function showSections() {
  ['statsSection','chartsSection','examFilterSection'].forEach(id =>
    document.getElementById(id).classList.remove('hidden'));
}

/* ─── Grade Distribution Panel ──────────────────────── */
const GRADE_BANDS = [
  { label: 'ممتاز',    range: '90 – 100', min: 90, max: 100, color: '#10b981', bg: '#d1fae5', text: '#065f46' },
  { label: 'جيد جداً', range: '80 – 89',  min: 80, max: 89,  color: '#3b82f6', bg: '#dbeafe', text: '#1e40af' },
  { label: 'جيد',      range: '70 – 79',  min: 70, max: 79,  color: '#f59e0b', bg: '#fef3c7', text: '#92400e' },
  { label: 'مقبول',    range: '60 – 69',  min: 60, max: 69,  color: '#f97316', bg: '#ffedd5', text: '#9a3412' },
  { label: 'ضعيف',     range: 'أقل من 60', min: 0,  max: 59,  color: '#ef4444', bg: '#fee2e2', text: '#991b1b' },
];

function updateGradeDist(scores, title) {
  if (!scores || !scores.length) return;

  const total = scores.length;
  document.getElementById('gradeDistSection').classList.remove('hidden');
  document.getElementById('gradeDistTitle').textContent = title || '';
  document.getElementById('gradeDistTotal').textContent = `${total} طالب`;

  document.getElementById('gradeDistRows').innerHTML = GRADE_BANDS.map(band => {
    const count = scores.filter(s => s >= band.min && s <= band.max).length;
    const pct   = total > 0 ? ((count / total) * 100).toFixed(1) : '0.0';
    const barW  = total > 0 ? Math.round((count / total) * 100) : 0;

    return `
    <div class="flex items-center gap-3">
      <div class="w-20 flex-shrink-0 text-right">
        <span class="text-xs font-bold px-2 py-0.5 rounded-full" style="background:${band.bg};color:${band.text}">${band.label}</span>
      </div>
      <div class="w-16 flex-shrink-0 text-xs text-slate-400 text-center">${band.range}</div>
      <div class="flex-1 bg-slate-100 rounded-full h-4 overflow-hidden">
        <div class="h-4 rounded-full transition-all duration-500" style="width:${barW}%;background:${band.color}"></div>
      </div>
      <div class="w-10 text-sm font-extrabold text-slate-700 text-center">${count}</div>
      <div class="w-14 text-sm font-semibold text-slate-400 text-left">${pct}%</div>
    </div>`;
  }).join('');
}

/* ─── Score Helpers ─────────────────────────────────── */
function getScores(data, examKey) {
  return examKey === 'all'
    ? data.map(r => avg([r['فتري1'], r['فتري2'], r['نهائي']]))
    : data.map(r => r[examKey]);
}

function avg(arr) {
  const v = arr.filter(x => !isNaN(x));
  return v.length ? v.reduce((a, b) => a + b, 0) / v.length : 0;
}

function getGrade(s) { return GRADES.find(g => s >= g.min) || GRADES[GRADES.length-1]; }

/* ─── Stats ─────────────────────────────────────────── */
function updateStats(data, scores) {
  if (!scores.length) return;
  const mean   = avg(scores);
  const passed = scores.filter(s => s >= 60).length;
  const maxV   = Math.max(...scores), minV = Math.min(...scores);
  const maxIdx = scores.indexOf(maxV), minIdx = scores.indexOf(minV);

  document.getElementById('statAvg').textContent  = mean.toFixed(1);
  document.getElementById('statPass').textContent = `${((passed/scores.length)*100).toFixed(0)}%`;
  document.getElementById('statPassDetail').textContent = `${passed} من ${scores.length} طالب`;
  document.getElementById('statMax').textContent  = maxV.toFixed(1);
  document.getElementById('statMin').textContent  = minV.toFixed(1);
  document.getElementById('statMaxName').textContent = data[maxIdx]?.name || '—';
  document.getElementById('statMinName').textContent = data[minIdx]?.name || '—';
}

/* ─── Bar Chart ─────────────────────────────────────── */
function updateBarChart(data, examKey) {
  const classes  = [...new Set(data.map(r => r.class))].sort((a,b) => +a-+b || a.localeCompare(b,'ar'));
  const averages = classes.map(cls => parseFloat(avg(getScores(data.filter(r=>r.class===cls), examKey)).toFixed(1)));
  const color    = STAGE_COLORS[currentStage];

  const ctx = document.getElementById('barChart').getContext('2d');
  if (barChartInst) barChartInst.destroy();
  barChartInst = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: classes.map(c => `الفصل ${c}`),
      datasets: [{ label: 'المتوسط', data: averages,
        backgroundColor: classes.map((_,i) => `${color}${['dd','bb','99','77'][i%4]}`),
        borderRadius: 8, borderSkipped: false }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => ` ${c.parsed.y} درجة` } } },
      scales: {
        y: { beginAtZero: false, min: 0, max: 100, grid: { color:'#f1f5f9' }, ticks: { font: { family:'Tajawal', size:11 } } },
        x: { grid: { display: false }, ticks: { font: { family:'Tajawal', size:12 } } },
      },
    },
  });
}

/* ─── Pie Chart ─────────────────────────────────────── */
function updatePieChart(scores) {
  const buckets = GRADES.map((g, i) => {
    const upper = i === 0 ? Infinity : GRADES[i-1].min;
    return scores.filter(s => s >= g.min && s < upper).length;
  });

  const ctx = document.getElementById('pieChart').getContext('2d');
  if (pieChartInst) pieChartInst.destroy();
  pieChartInst = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: GRADES.map(g => g.label),
      datasets: [{ data: buckets, backgroundColor: GRADES.map(g=>g.color), borderWidth: 2, borderColor:'#fff' }],
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '62%',
      plugins: { legend: { display:false }, tooltip: { callbacks: { label: c => ` ${c.label}: ${c.parsed} طالب` } } },
    },
  });

  document.getElementById('gradeLegend').innerHTML =
    GRADES.map((g,i) => `<div class="flex items-center gap-1.5">
      <span class="w-3 h-3 rounded-full flex-shrink-0" style="background:${g.color}"></span>
      <span>${g.label} (${buckets[i]})</span></div>`).join('');
}

/* ─── Line Chart ─────────────────────────────────────── */
function updateLineChart(stageData) {
  const classes = [...new Set(stageData.map(r => r.class))].sort((a,b) => +a-+b || a.localeCompare(b,'ar'));
  const base    = STAGE_COLORS[currentStage];
  const palette = classes.map((_,i) => `hsl(${(currentStage==='ثانوية'?210:145) + i*35}, 65%, ${48+i*4}%)`);

  if (currentClass !== 'all') {
    const filtered = classes.filter(c => c === currentClass);
    const datasets = filtered.map((cls, i) => ({
      label: `الفصل ${cls}`,
      data:  EXAM_COLS.map(e => parseFloat(avg(stageData.filter(r=>r.class===cls).map(r=>r[e])).toFixed(1))),
      borderColor: palette[classes.indexOf(cls)],
      backgroundColor: palette[classes.indexOf(cls)] + '22',
      borderWidth: 2.5, pointRadius: 5, tension: 0.35, fill: false,
    }));
    renderLineChart(datasets);
  } else {
    const datasets = classes.map((cls, i) => ({
      label: `الفصل ${cls}`,
      data:  EXAM_COLS.map(e => parseFloat(avg(stageData.filter(r=>r.class===cls).map(r=>r[e])).toFixed(1))),
      borderColor: palette[i],
      backgroundColor: palette[i] + '22',
      borderWidth: 2.5, pointRadius: 5, tension: 0.35, fill: false,
    }));
    renderLineChart(datasets);
  }
}

function renderLineChart(datasets) {
  const ctx = document.getElementById('lineChart').getContext('2d');
  if (lineChartInst) lineChartInst.destroy();
  lineChartInst = new Chart(ctx, {
    type: 'line',
    data: { labels: EXAM_COLS.map(e => EXAM_LABELS[e]), datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position:'bottom', labels: { font:{family:'Tajawal',size:12}, boxWidth:12, padding:14 } },
        tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${c.parsed.y} درجة` } },
      },
      scales: {
        y: { min:0, max:100, grid:{color:'#f1f5f9'}, ticks:{font:{family:'Tajawal',size:11}} },
        x: { grid:{display:false}, ticks:{font:{family:'Tajawal',size:12}} },
      },
    },
  });
}

/* ─── Add Class Modal ───────────────────────────────── */
function openAddClassModal() {
  document.getElementById('newClassStage').value = currentStage;
  document.getElementById('newClassName').value  = '';
  const m = document.getElementById('addClassModal');
  m.classList.remove('hidden');
  m.classList.add('modal-open');
  setTimeout(() => document.getElementById('newClassName').focus(), 50);
}

function closeAddClassModal() {
  const m = document.getElementById('addClassModal');
  m.classList.add('hidden');
  m.classList.remove('modal-open');
}

function confirmAddClass() {
  const stage = document.getElementById('newClassStage').value.trim();
  const name  = document.getElementById('newClassName').value.trim();
  if (!name) { alert('أدخل رقم أو اسم الفصل.'); return; }

  const stageData = rawData.filter(r => r.stage === stage);
  const existing  = [...new Set(stageData.map(r => r.class)), ...extraClasses[stage]];
  if (existing.includes(name)) { alert('هذا الفصل موجود مسبقاً.'); return; }

  extraClasses[stage].push(name);
  closeAddClassModal();
  if (stage === currentStage) { renderClassChips(); applyFilters(); }
}

document.getElementById('addClassModal').addEventListener('click', function(e) {
  if (e.target === this) closeAddClassModal();
});

document.getElementById('newClassName').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') confirmAddClass();
});

document.getElementById('editClassModal').addEventListener('click', function(e) {
  if (e.target === this) closeEditClassModal();
});

document.getElementById('editClassName').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') confirmEditClass();
});

/* ─── Print ─────────────────────────────────────────── */
function printReport() {
  const stageData = rawData.filter(r => r.stage === currentStage);
  const data = currentClass !== 'all' ? stageData.filter(r => r.class === currentClass) : stageData;
  const scores = getScores(data, currentExam);

  const label = currentClass === 'all'
    ? `المرحلة ${currentStage} — جميع الفصول`
    : `المرحلة ${currentStage} — الفصل ${currentClass}`;
  const examLabel = currentExam === 'all' ? 'جميع الاختبارات' : EXAM_LABELS[currentExam];

  document.getElementById('printSubtitle').textContent = `${label}  |  ${examLabel}`;

  // Stats block
  const mean   = avg(scores);
  const passed = scores.filter(s => s >= 60).length;
  const maxV   = Math.max(...scores), minV = Math.min(...scores);
  const maxIdx = scores.indexOf(maxV), minIdx = scores.indexOf(minV);
  const color  = STAGE_COLORS[currentStage];

  document.getElementById('printStats').innerHTML = [
    { title: 'المتوسط العام', value: mean.toFixed(1), sub: 'من 100 درجة' },
    { title: 'نسبة النجاح',  value: `${((passed/scores.length)*100).toFixed(0)}%`, sub: `${passed} من ${scores.length} طالب` },
    { title: 'أعلى درجة',   value: maxV.toFixed(1), sub: data[maxIdx]?.name || '' },
    { title: 'أقل درجة',    value: minV.toFixed(1), sub: data[minIdx]?.name || '' },
  ].map(s => `
    <div style="border:1px solid #e2e8f0; border-radius:12px; padding:14px; text-align:right;">
      <div style="font-size:11px; color:#64748b; margin-bottom:6px;">${s.title}</div>
      <div style="font-size:26px; font-weight:900; color:#1e293b;">${s.value}</div>
      <div style="font-size:11px; color:#94a3b8; margin-top:3px;">${s.sub}</div>
    </div>`).join('');

  // Destroy old print charts
  ['printBarChart','printPieChart','printLineChart'].forEach(id => {
    const existing = Chart.getChart(id);
    if (existing) existing.destroy();
  });

  // Bar
  const classes  = [...new Set(data.map(r=>r.class))].sort((a,b)=>+a-+b||a.localeCompare(b,'ar'));
  const averages = classes.map(cls => parseFloat(avg(getScores(data.filter(r=>r.class===cls), currentExam)).toFixed(1)));
  new Chart(document.getElementById('printBarChart').getContext('2d'), {
    type:'bar',
    data:{ labels: classes.map(c=>`الفصل ${c}`),
      datasets:[{label:'المتوسط', data:averages, backgroundColor: color+'cc', borderRadius:6, borderSkipped:false}]},
    options:{ responsive:true, plugins:{legend:{display:false}}, scales:{y:{min:0,max:100}} },
  });

  // Pie
  const buckets = GRADES.map((g,i)=>{const upper=i===0?Infinity:GRADES[i-1].min; return scores.filter(s=>s>=g.min&&s<upper).length;});
  new Chart(document.getElementById('printPieChart').getContext('2d'), {
    type:'doughnut',
    data:{ labels:GRADES.map(g=>g.label), datasets:[{data:buckets, backgroundColor:GRADES.map(g=>g.color), borderWidth:2, borderColor:'#fff'}]},
    options:{ responsive:true, cutout:'55%', plugins:{legend:{position:'bottom', labels:{font:{size:10}}}} },
  });

  // Line
  const sData = rawData.filter(r=>r.stage===currentStage);
  const lClasses = currentClass!=='all' ? [currentClass] : [...new Set(sData.map(r=>r.class))].sort((a,b)=>+a-+b||a.localeCompare(b,'ar'));
  const palette  = lClasses.map((_,i)=>`hsl(${(currentStage==='ثانوية'?210:145)+i*35}, 65%, ${48+i*4}%)`);
  new Chart(document.getElementById('printLineChart').getContext('2d'), {
    type:'line',
    data:{ labels: EXAM_COLS.map(e=>EXAM_LABELS[e]),
      datasets: lClasses.map((cls,i)=>({
        label:`الفصل ${cls}`,
        data: EXAM_COLS.map(e=>parseFloat(avg(sData.filter(r=>r.class===cls).map(r=>r[e])).toFixed(1))),
        borderColor:palette[i], borderWidth:2, pointRadius:4, tension:0.35, fill:false,
      }))},
    options:{ responsive:true, plugins:{legend:{position:'bottom', labels:{font:{size:10}}}}, scales:{y:{min:0,max:100}} },
  });

  setTimeout(() => window.print(), 300);
}
