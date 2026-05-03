/* ══════════════════════════════════════════
   app.js — مولّد نماذج الامتحانات
   Dependencies: xlsx.js, jszip.js
══════════════════════════════════════════ */

// ══════════════════════════════════════════
// STATE
// ══════════════════════════════════════════
let questions       = [];
let models          = [];
let currentModelIdx = 0;

let colMap = { q:'', a:'', b:'', c:'', d:'', ans:'', topic:'', diff:'' };

let opts = {
  shuffleQ:    true,
  shuffleA:    true,
  unique:      true,
  header:      true,
  student:     true,
  modelNum:    true,
  answerKey:   true,
  answerTable: true
};

let design = {
  headerBg:   '#1e3a8a',
  headerText: '#ffffff',
  fontSize:   14,
  layoutCols: 2
};

let examCfg = {};

// ══════════════════════════════════════════
// NAVIGATION
// ══════════════════════════════════════════
function goPanel(n) {
  if (n > 0 && questions.length === 0) {
    showToast('يرجى استيراد الأسئلة أولاً', 'error');
    return;
  }
  document.querySelectorAll('.panel').forEach((p, i) => p.classList.toggle('active', i === n));
  document.querySelectorAll('.step-btn').forEach((b, i) => b.classList.toggle('active', i === n));
  updatePreview();
}

// ══════════════════════════════════════════
// SIDEBAR TOGGLES
// ══════════════════════════════════════════
function toggleOpt(el) {
  el.classList.toggle('on');
  const k = el.id.replace('tgl-', '');
  const map = {
    'shuffle-q':   'shuffleQ',
    'shuffle-a':   'shuffleA',
    'unique':      'unique',
    'header':      'header',
    'student':     'student',
    'modelnum':    'modelNum',
    'answerkey':   'answerKey',
    'answertable': 'answerTable'
  };
  if (map[k] !== undefined) opts[map[k]] = el.classList.contains('on');
}

// ══════════════════════════════════════════
// FILE IMPORT
// ══════════════════════════════════════════
function initDropZone() {
  const dropZone = document.getElementById('dropZone');
  dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('drag'); });
  dropZone.addEventListener('dragleave', ()  => dropZone.classList.remove('drag'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag');
    handleFile({ target: { files: e.dataTransfer.files } });
  });
}

function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb  = XLSX.read(e.target.result, { type: 'array' });
      const ws  = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
      processData(raw);
    } catch (err) {
      showToast('خطأ في قراءة الملف: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function processData(raw) {
  if (raw.length < 2) { showToast('الملف فارغ', 'error'); return; }
  const headers = raw[0].map(h => String(h || '').trim());
  autoMap(headers);
  questions = raw.slice(1)
    .filter(r => r.some(c => c !== undefined && c !== ''))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i] !== undefined ? String(r[i]).trim() : '');
      return obj;
    });
  updateStats(headers);
  renderPreviewTable(headers);
  showToast('تم استيراد ' + questions.length + ' سؤال', 'success');
}

// ══════════════════════════════════════════
// COLUMN AUTO-MAPPING
// ══════════════════════════════════════════
function autoMap(headers) {
  const lh = headers.map(h => String(h || '').trim().toLowerCase());

  const findExact = targets => {
    for (const t of targets) {
      const i = lh.findIndex(h => h === t.toLowerCase());
      if (i !== -1) return headers[i];
    }
    return '';
  };
  const findContains = targets => {
    for (const t of targets) {
      const i = lh.findIndex(h => h.includes(t.toLowerCase()));
      if (i !== -1) return headers[i];
    }
    return '';
  };

  colMap.q     = findExact(['Question','question','سؤال','نص السؤال'])   || findContains(['question','سؤال','نص']);
  colMap.a     = findExact(['A','a','أ','ا','Choice A','Option A'])        || findContains(['choice a','option a','خيار أ']);
  colMap.b     = findExact(['B','b','ب','Choice B','Option B'])            || findContains(['choice b','option b','خيار ب']);
  colMap.c     = findExact(['C','c','ج','Choice C','Option C'])            || findContains(['choice c','option c','خيار ج']);
  colMap.d     = findExact(['D','d','د','Choice D','Option D'])            || findContains(['choice d','option d','خيار د']);
  colMap.ans   = findExact(['Answer','answer','Correct','correct','الإجابة','إجابة','الإجابة الصحيحة']) || findContains(['answer','correct','إجابة','صحيح']);
  colMap.topic = findExact(['Topic','topic','Unit','unit','محور','وحدة','فصل']) || findContains(['topic','unit','محور','وحدة']);
  colMap.diff  = findExact(['Difficulty','difficulty','Level','level','صعوبة'])  || findContains(['difficulty','level','صعوبة']);

  if (!colMap.q   && headers[0]) colMap.q   = headers[0];
  if (!colMap.a   && headers[1]) colMap.a   = headers[1];
  if (!colMap.b   && headers[2]) colMap.b   = headers[2];
  if (!colMap.c   && headers[3]) colMap.c   = headers[3];
  if (!colMap.d   && headers[4]) colMap.d   = headers[4];
  if (!colMap.ans && headers[5]) colMap.ans = headers[5];
}

// ══════════════════════════════════════════
// STATS & PREVIEW TABLE
// ══════════════════════════════════════════
function updateStats(headers) {
  const topics = new Set(questions.map(q => q[colMap.topic]).filter(Boolean));
  document.getElementById('stat-total').textContent  = questions.length;
  document.getElementById('stat-topics').textContent = topics.size || '—';
  document.getElementById('stat-cols').textContent   = headers.length;
  document.getElementById('stat-ready').textContent  = '✅';
  document.getElementById('sb-total').textContent    = questions.length;
  document.getElementById('sb-topics').textContent   = topics.size || 0;
  document.getElementById('sb-avg').textContent      = [colMap.a,colMap.b,colMap.c,colMap.d].filter(Boolean).length;
}

function renderPreviewTable(headers) {
  const showCols = headers.slice(0, 6);
  document.getElementById('table-head').innerHTML =
    '<th>#</th>' + showCols.map(h => `<th>${h}</th>`).join('');
  document.getElementById('table-body').innerHTML =
    questions.slice(0, 8).map((q, i) =>
      `<tr><td>${i+1}</td>${showCols.map(h =>
        `<td>${String(q[h]||'').slice(0,60)}</td>`).join('')}</tr>`
    ).join('')
    + (questions.length > 8
      ? `<tr><td colspan="${showCols.length+1}" style="text-align:center;color:var(--text3)">
           ... و ${questions.length-8} سؤال آخر</td></tr>` : '');
  document.getElementById('preview-table').style.display = 'block';
}

// ══════════════════════════════════════════
// COLOR PICKER
// ══════════════════════════════════════════
function initColorPickers() {
  document.getElementById('headerBg').addEventListener('input', e => {
    design.headerBg = e.target.value;
    document.getElementById('preview-header').style.background = e.target.value;
  });
  document.getElementById('headerText').addEventListener('input', e => {
    design.headerText = e.target.value;
    document.getElementById('preview-header').style.color = e.target.value;
  });
}

function pickColor(swatch, inputId) {
  const color = swatch.dataset.color;
  swatch.parentElement.querySelectorAll('.color-swatch').forEach(s => s.classList.remove('selected'));
  swatch.classList.add('selected');
  document.getElementById(inputId).value = color;
  if (inputId === 'headerBg')   { design.headerBg   = color; document.getElementById('preview-header').style.background = color; }
  if (inputId === 'headerText') { design.headerText = color; document.getElementById('preview-header').style.color      = color; }
}

function updatePreview() {
  const g = id => document.getElementById(id);
  if (g('prev-institution')) g('prev-institution').textContent = g('institution')?.value || '';
  if (g('prev-subject'))     g('prev-subject').textContent     = 'امتحان مادة: ' + (g('subjectName')?.value||'');
  const ph = g('preview-header');
  if (ph) {
    const last = ph.querySelector('div:last-child');
    if (last) last.textContent = `الزمن: ${g('examDuration')?.value||''} دقيقة | الدرجة: ${g('totalGrade')?.value||''}`;
  }
  design.fontSize   = parseInt(g('fontSize')?.value   || 14);
  design.layoutCols = parseInt(g('layoutCols')?.value || 2);
}

// ══════════════════════════════════════════
// SHUFFLE — Fisher-Yates
// ══════════════════════════════════════════
function shuffle(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

// ══════════════════════════════════════════
// DEDUPLICATE choices — remove blank/duplicate option text
// ══════════════════════════════════════════
function deduplicateChoices(rawChoices) {
  const seen = new Set();
  return rawChoices.filter(c => {
    const val = (c.val || '').trim();
    if (!val) return false;
    if (seen.has(val)) return false;
    seen.add(val);
    return true;
  });
}

// Map English answer letter → Arabic label
function mapEnToAr(letter) {
  return { A:'أ', B:'ب', C:'ج', D:'د', a:'أ', b:'ب', c:'ج', d:'د' }[letter] || letter;
}

// ══════════════════════════════════════════
// GENERATE MODELS
// Guarantees: no duplicate questions per model, no duplicate choices per question
// ══════════════════════════════════════════
function generateModels() {
  if (!questions.length) { showToast('لا توجد أسئلة للتوليد', 'error'); return; }

  const n      = parseInt(document.getElementById('numModels').value)  || 4;
  const qCount = parseInt(document.getElementById('qPerModel').value)  || 20;

  if (qCount > questions.length) {
    showToast(`عدد الأسئلة المطلوب (${qCount}) أكبر من المتاح (${questions.length})`, 'error');
    return;
  }

  examCfg = {
    subject:      document.getElementById('subjectName').value,
    institution:  document.getElementById('institution').value,
    duration:     document.getElementById('examDuration').value,
    grade:        document.getElementById('totalGrade').value,
    classLevel:   document.getElementById('classLevel').value,
    date:         document.getElementById('examDate').value,
    instructions: document.getElementById('instructions').value,
    qCount, n
  };

  const btn = document.getElementById('gen-btn');
  btn.classList.add('generating');

  setTimeout(() => {
    models = [];
    const modelLetters = 'أبجدهوزحطيكلمنسعفصقرشتثخذضظغ'.split('');

    // Track globally used question indices for unique mode
    const globalUsedIndices = new Set();

    for (let m = 0; m < n; m++) {
      // Build candidate pool (by index to guarantee uniqueness)
      let pool = questions.map((q, idx) => ({ q, idx }));

      if (opts.unique) {
        const remaining = pool.filter(({ idx }) => !globalUsedIndices.has(idx));
        if (remaining.length >= qCount) pool = remaining;
        // else: wrap around — reuse all questions
      }

      // Shuffle and pick exactly qCount — NO duplicates within a model (each idx appears once)
      const selected = shuffle(pool).slice(0, qCount);
      selected.forEach(({ idx }) => globalUsedIndices.add(idx));

      // Process each question
      const processedQ = selected.map(({ q }) => {
        const rawChoices = [
          { label: 'أ', val: (q[colMap.a] || '').trim() },
          { label: 'ب', val: (q[colMap.b] || '').trim() },
          { label: 'ج', val: (q[colMap.c] || '').trim() },
          { label: 'د', val: (q[colMap.d] || '').trim() }
        ];

        // Remove blank + duplicate choice texts
        const uniqueChoices = deduplicateChoices(rawChoices);

        // Shuffle choices
        const finalChoices = opts.shuffleA ? shuffle(uniqueChoices) : uniqueChoices;

        // Resolve correct answer after shuffling
        const correctRaw = (q[colMap.ans] || '').trim();
        const correctChoice = finalChoices.find(c =>
          c.val === correctRaw ||
          c.label === correctRaw ||
          c.label === mapEnToAr(correctRaw)
        );
        const correctLabel = correctChoice ? correctChoice.label : correctRaw;

        return {
          text:         (q[colMap.q] || '').trim(),
          choices:      finalChoices,
          correctLabel,
          topic:        (q[colMap.topic] || '').trim(),
          diff:         (q[colMap.diff]  || '').trim()
        };
      });

      models.push({
        name:      'نموذج ' + modelLetters[m],
        letter:    modelLetters[m],
        questions: opts.shuffleQ ? shuffle(processedQ) : processedQ
      });
    }

    btn.classList.remove('generating');
    renderModelTabs();
    goPanel(3);
    document.getElementById('gen-summary').textContent =
      `${n} نماذج × ${qCount} سؤال = ${n * qCount} سؤال إجمالي`;
    showToast(`تم توليد ${n} نماذج بنجاح`, 'success');
  }, 800);
}

// ══════════════════════════════════════════
// RENDER MODEL TABS
// ══════════════════════════════════════════
function renderModelTabs() {
  const tabs = document.getElementById('model-tabs');
  tabs.innerHTML =
    models.map((m, i) =>
      `<button class="exam-tab ${i===0?'active':''}" onclick="showModel(${i})">${m.name}</button>`
    ).join('')
    + `<button class="exam-tab" onclick="showAnswerKeys()">نماذج الإجابات</button>`;
  showModel(0);
}

function showModel(idx) {
  currentModelIdx = idx;
  document.querySelectorAll('.exam-tab').forEach((t,i) => t.classList.toggle('active', i===idx));
  document.getElementById('answer-key-view').style.display = 'none';
  document.getElementById('model-view').style.display      = 'block';
  document.getElementById('model-view').innerHTML          = renderExamPaper(models[idx], idx);
}

function showAnswerKeys() {
  document.querySelectorAll('.exam-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.exam-tab').item(models.length).classList.add('active');
  document.getElementById('model-view').style.display      = 'none';
  document.getElementById('answer-key-view').style.display = 'block';
  renderAllAnswerKeys();
}

// ══════════════════════════════════════════
// RENDER EXAM PAPER — in-app preview
// ══════════════════════════════════════════
function renderExamPaper(model, idx) {
  const fs  = design.fontSize;
  const hBg = design.headerBg;

  const header = opts.header ? `
    <div style="background:${hBg};color:${design.headerText};
                padding:1.5rem;text-align:center;margin:-2.5rem -2.5rem 1.5rem;">
      <div style="font-size:${fs+2}px;font-weight:700;margin-bottom:4px;">${examCfg.institution}</div>
      <div style="font-size:${fs+8}px;font-weight:900;font-family:'Amiri',serif;">امتحان مادة: ${examCfg.subject}</div>
      <div style="font-size:${fs-1}px;opacity:.85;margin-top:4px;">الزمن: ${examCfg.duration} دق | الدرجة: ${examCfg.grade}</div>
      ${opts.modelNum ? `<div style="margin-top:6px;display:inline-block;background:rgba(255,255,255,.25);
        padding:3px 16px;border-radius:20px;font-size:${fs+1}px;font-weight:700;">النموذج: ${model.letter}</div>` : ''}
    </div>` : '';

  const student = opts.student ? `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:1.5rem;font-size:${fs-1}px;">
      <div style="border-bottom:1px solid #2d3748;padding-bottom:4px;color:#4a5568;font-weight:500;">اسم الطالب: ________________</div>
      <div style="border-bottom:1px solid #2d3748;padding-bottom:4px;color:#4a5568;font-weight:500;">رقم الجلوس: ____________</div>
      <div style="border-bottom:1px solid #2d3748;padding-bottom:4px;color:#4a5568;font-weight:500;">الشعبة / الفصل: __________</div>
      <div style="border-bottom:1px solid #2d3748;padding-bottom:4px;color:#4a5568;font-weight:500;">التاريخ: ${examCfg.date||'___________'}</div>
    </div>` : '';

  const instructions = examCfg.instructions
    ? `<div style="background:#f8f9ff;border:1px solid #d0d8f0;border-radius:8px;
                   padding:10px 14px;margin-bottom:1.5rem;font-size:${fs-1}px;color:#4a5568;line-height:1.8;">
         <strong>التعليمات:</strong> ${examCfg.instructions}</div>` : '';

  const qItems = model.questions.map((item, i) => {
    const optHtml = item.choices.map(c => `
      <div style="display:flex;gap:6px;align-items:flex-start;font-size:${fs-1}px;color:#4a5568;margin-bottom:4px;">
        <span style="font-weight:700;color:#2d3748;flex-shrink:0;">${c.label})</span>
        <span>${c.val}</span>
      </div>`).join('');
    return `
      <div style="display:block;width:100%;margin:0 0 1.4rem 0;padding-bottom:1rem;
                  border-bottom:1px solid #e0e4ef;page-break-inside:avoid;">
        <div style="font-size:${fs}px;line-height:1.9;margin-bottom:8px;color:#1a1a2e;font-weight:500;">
          <strong style="color:${hBg};">${i+1}.</strong> ${item.text}
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px 16px;padding-right:16px;">${optHtml}</div>
      </div>`;
  }).join('');

  const answerTable = opts.answerTable ? buildAnswerFillTable(model.questions.length, hBg, fs) : '';

  return `
    <div class="exam-paper" id="paper-${idx}">
      ${header}${student}${instructions}
      <div style="display:block;width:100%;">${qItems}</div>
      ${answerTable}
    </div>`;
}

// ══════════════════════════════════════════
// ANSWER FILL TABLE — blank table at end of question paper
// ══════════════════════════════════════════
function buildAnswerFillTable(count, hBg, fs) {
  const cells = Array.from({ length: count }, (_, i) => `
    <td style="text-align:center;border:1px solid #c8d0e0;padding:8px 4px;min-width:44px;">
      <div style="font-size:9px;color:#888;margin-bottom:4px;">${i+1}</div>
      <div style="height:22px;border-bottom:1px solid #555;"></div>
    </td>`);
  const rows = [];
  for (let r = 0; r < cells.length; r += 10)
    rows.push(`<tr>${cells.slice(r, r+10).join('')}</tr>`);
  return `
    <div style="margin-top:2rem;page-break-inside:avoid;">
      <div style="background:${hBg};color:#fff;padding:6px 14px;font-size:${fs-1}px;
                  font-weight:700;border-radius:6px 6px 0 0;display:inline-block;">جدول الإجابات</div>
      <table style="width:100%;border-collapse:collapse;border:1px solid #c8d0e0;">${rows.join('')}</table>
    </div>`;
}

// ══════════════════════════════════════════
// RENDER ANSWER KEYS — in-app
// ══════════════════════════════════════════
function renderAllAnswerKeys() {
  document.getElementById('answer-keys-content').innerHTML = models.map(m => {
    const grid = m.questions.map((q, i) => `
      <div class="ans-item">
        <div class="ans-qnum">${i+1}</div>
        <div class="ans-val">${q.correctLabel}</div>
      </div>`).join('');
    return `
      <div style="margin-bottom:2rem;">
        <div style="font-size:15px;font-weight:700;margin-bottom:.75rem;color:var(--text);
                    display:flex;align-items:center;gap:8px;">
          <span style="background:${design.headerBg};color:${design.headerText};
                        padding:4px 14px;border-radius:20px;font-size:13px;">${m.name}</span>
          ${examCfg.subject} — ${m.questions.length} سؤال
        </div>
        <div style="display:grid;grid-template-columns:repeat(10,1fr);gap:8px;">${grid}</div>
      </div>`;
  }).join('');
}

// ══════════════════════════════════════════
// SHARED CSS FOR EXPORTED FILES
// ══════════════════════════════════════════
function exportPageCss(extra) {
  return `
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&family=Amiri:wght@400;700&display=swap');
    @page { size:A4; margin:18mm 20mm 20mm 20mm; }
    @page { @bottom-center {
      content: "صفحة " counter(page) " من " counter(pages);
      font-family:'Cairo',sans-serif; font-size:9pt; color:#666;
    }}
    * { box-sizing:border-box; margin:0; padding:0; }
    body { font-family:'Cairo',sans-serif; /* ══════════════════════════════════════════
   app.js — مولّد نماذج الامتحانات
   Dependencies: xlsx.js, jszip.js
══════════════════════════════════════════ */

// ══════════════════════════════════════════
// STATE
// ══════════════════════════════════════════
let questions       = [];
let models          = [];
let currentModelIdx = 0;

let colMap = { q:'', a:'', b:'', c:'', d:'', ans:'', topic:'', diff:'' };

let opts = {
  shuffleQ:    true,
  shuffleA:    true,
  unique:      true,
  header:      true,
  student:     true,
  modelNum:    true,
  answerKey:   true,
  answerTable: true
};

let design = {
  headerBg:   '#1e3a8a',
  headerText: '#ffffff',
  fontSize:   14,
  layoutCols: 2
};

let examCfg = {};

// ══════════════════════════════════════════
// NAVIGATION
// ══════════════════════════════════════════
function goPanel(n) {
  if (n > 0 && questions.length === 0) {
    showToast('يرجى استيراد الأسئلة أولاً', 'error');
    return;
  }
  document.querySelectorAll('.panel').forEach((p, i) => p.classList.toggle('active', i === n));
  document.querySelectorAll('.step-btn').forEach((b, i) => b.classList.toggle('active', i === n));
  updatePreview();
}

// ══════════════════════════════════════════
// SIDEBAR TOGGLES
// ══════════════════════════════════════════
function toggleOpt(el) {
  el.classList.toggle('on');
  const k = el.id.replace('tgl-', '');
  const map = {
    'shuffle-q':   'shuffleQ',
    'shuffle-a':   'shuffleA',
    'unique':      'unique',
    'header':      'header',
    'student':     'student',
    'modelnum':    'modelNum',
    'answerkey':   'answerKey',
    'answertable': 'answerTable'
  };
  if (map[k] !== undefined) opts[map[k]] = el.classList.contains('on');
}

// ══════════════════════════════════════════
// FILE IMPORT
// ══════════════════════════════════════════
function initDropZone() {
  const dropZone = document.getElementById('dropZone');
  dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('drag'); });
  dropZone.addEventListener('dragleave', ()  => dropZone.classList.remove('drag'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag');
    handleFile({ target: { files: e.dataTransfer.files } });
  });
}

function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb  = XLSX.read(e.target.result, { type: 'array' });
      const ws  = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
      processData(raw);
    } catch (err) {
      showToast('خطأ في قراءة الملف: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function processData(raw) {
  if (raw.length < 2) { showToast('الملف فارغ', 'error'); return; }
  const headers = raw[0].map(h => String(h || '').trim());
  autoMap(headers);
  questions = raw.slice(1)
    .filter(r => r.some(c => c !== undefined && c !== ''))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i] !== undefined ? String(r[i]).trim() : '');
      return obj;
    });
  updateStats(headers);
  renderPreviewTable(headers);
  showToast('تم استيراد ' + questions.length + ' سؤال', 'success');
}

// ══════════════════════════════════════════
// COLUMN AUTO-MAPPING
// ══════════════════════════════════════════
function autoMap(headers) {
  const lh = headers.map(h => String(h || '').trim().toLowerCase());

  const findExact = targets => {
    for (const t of targets) {
      const i = lh.findIndex(h => h === t.toLowerCase());
      if (i !== -1) return headers[i];
    }
    return '';
  };
  const findContains = targets => {
    for (const t of targets) {
      const i = lh.findIndex(h => h.includes(t.toLowerCase()));
      if (i !== -1) return headers[i];
    }
    return '';
  };

  colMap.q     = findExact(['Question','question','سؤال','نص السؤال'])   || findContains(['question','سؤال','نص']);
  colMap.a     = findExact(['A','a','أ','ا','Choice A','Option A'])        || findContains(['choice a','option a','خيار أ']);
  colMap.b     = findExact(['B','b','ب','Choice B','Option B'])            || findContains(['choice b','option b','خيار ب']);
  colMap.c     = findExact(['C','c','ج','Choice C','Option C'])            || findContains(['choice c','option c','خيار ج']);
  colMap.d     = findExact(['D','d','د','Choice D','Option D'])            || findContains(['choice d','option d','خيار د']);
  colMap.ans   = findExact(['Answer','answer','Correct','correct','الإجابة','إجابة','الإجابة الصحيحة']) || findContains(['answer','correct','إجابة','صحيح']);
  colMap.topic = findExact(['Topic','topic','Unit','unit','محور','وحدة','فصل']) || findContains(['topic','unit','محور','وحدة']);
  colMap.diff  = findExact(['Difficulty','difficulty','Level','level','صعوبة'])  || findContains(['difficulty','level','صعوبة']);

  if (!colMap.q   && headers[0]) colMap.q   = headers[0];
  if (!colMap.a   && headers[1]) colMap.a   = headers[1];
  if (!colMap.b   && headers[2]) colMap.b   = headers[2];
  if (!colMap.c   && headers[3]) colMap.c   = headers[3];
  if (!colMap.d   && headers[4]) colMap.d   = headers[4];
  if (!colMap.ans && headers[5]) colMap.ans = headers[5];
}

// ══════════════════════════════════════════
// STATS & PREVIEW TABLE
// ══════════════════════════════════════════
function updateStats(headers) {
  const topics = new Set(questions.map(q => q[colMap.topic]).filter(Boolean));
  document.getElementById('stat-total').textContent  = questions.length;
  document.getElementById('stat-topics').textContent = topics.size || '—';
  document.getElementById('stat-cols').textContent   = headers.length;
  document.getElementById('stat-ready').textContent  = '✅';
  document.getElementById('sb-total').textContent    = questions.length;
  document.getElementById('sb-topics').textContent   = topics.size || 0;
  document.getElementById('sb-avg').textContent      = [colMap.a,colMap.b,colMap.c,colMap.d].filter(Boolean).length;
}

function renderPreviewTable(headers) {
  const showCols = headers.slice(0, 6);
  document.getElementById('table-head').innerHTML =
    '<th>#</th>' + showCols.map(h => `<th>${h}</th>`).join('');
  document.getElementById('table-body').innerHTML =
    questions.slice(0, 8).map((q, i) =>
      `<tr><td>${i+1}</td>${showCols.map(h =>
        `<td>${String(q[h]||'').slice(0,60)}</td>`).join('')}</tr>`
    ).join('')
    + (questions.length > 8
      ? `<tr><td colspan="${showCols.length+1}" style="text-align:center;color:var(--text3)">
           ... و ${questions.length-8} سؤال آخر</td></tr>` : '');
  document.getElementById('preview-table').style.display = 'block';
}

// ══════════════════════════════════════════
// COLOR PICKER
// ══════════════════════════════════════════
function initColorPickers() {
  document.getElementById('headerBg').addEventListener('input', e => {
    design.headerBg = e.target.value;
    document.getElementById('preview-header').style.background = e.target.value;
  });
  document.getElementById('headerText').addEventListener('input', e => {
    design.headerText = e.target.value;
    document.getElementById('preview-header').style.color = e.target.value;
  });
}

function pickColor(swatch, inputId) {
  const color = swatch.dataset.color;
  swatch.parentElement.querySelectorAll('.color-swatch').forEach(s => s.classList.remove('selected'));
  swatch.classList.add('selected');
  document.getElementById(inputId).value = color;
  if (inputId === 'headerBg')   { design.headerBg   = color; document.getElementById('preview-header').style.background = color; }
  if (inputId === 'headerText') { design.headerText = color; document.getElementById('preview-header').style.color      = color; }
}

function updatePreview() {
  const g = id => document.getElementById(id);
  if (g('prev-institution')) g('prev-institution').textContent = g('institution')?.value || '';
  if (g('prev-subject'))     g('prev-subject').textContent     = 'امتحان مادة: ' + (g('subjectName')?.value||'');
  const ph = g('preview-header');
  if (ph) {
    const last = ph.querySelector('div:last-child');
    if (last) last.textContent = `الزمن: ${g('examDuration')?.value||''} دقيقة | الدرجة: ${g('totalGrade')?.value||''}`;
  }
  design.fontSize   = parseInt(g('fontSize')?.value   || 14);
  design.layoutCols = parseInt(g('layoutCols')?.value || 2);
}

// ══════════════════════════════════════════
// SHUFFLE — Fisher-Yates
// ══════════════════════════════════════════
function shuffle(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

// ══════════════════════════════════════════
// DEDUPLICATE choices — remove blank/duplicate option text
// ══════════════════════════════════════════
function deduplicateChoices(rawChoices) {
  const seen = new Set();
  return rawChoices.filter(c => {
    const val = (c.val || '').trim();
    if (!val) return false;
    if (seen.has(val)) return false;
    seen.add(val);
    return true;
  });
}

// Map English answer letter → Arabic label
function mapEnToAr(letter) {
  return { A:'أ', B:'ب', C:'ج', D:'د', a:'أ', b:'ب', c:'ج', d:'د' }[letter] || letter;
}

// ══════════════════════════════════════════
// GENERATE MODELS
// Guarantees: no duplicate questions per model, no duplicate choices per question
// ══════════════════════════════════════════
function generateModels() {
  if (!questions.length) { showToast('لا توجد أسئلة للتوليد', 'error'); return; }

  const n      = parseInt(document.getElementById('numModels').value)  || 4;
  const qCount = parseInt(document.getElementById('qPerModel').value)  || 20;

  if (qCount > questions.length) {
    showToast(`عدد الأسئلة المطلوب (${qCount}) أكبر من المتاح (${questions.length})`, 'error');
    return;
  }

  examCfg = {
    subject:      document.getElementById('subjectName').value,
    institution:  document.getElementById('institution').value,
    duration:     document.getElementById('examDuration').value,
    grade:        document.getElementById('totalGrade').value,
    classLevel:   document.getElementById('classLevel').value,
    date:         document.getElementById('examDate').value,
    instructions: document.getElementById('instructions').value,
    qCount, n
  };

  const btn = document.getElementById('gen-btn');
  btn.classList.add('generating');

  setTimeout(() => {
    models = [];
    const modelLetters = 'أبجدهوزحطيكلمنسعفصقرشتثخذضظغ'.split('');

    // Track globally used question indices for unique mode
    const globalUsedIndices = new Set();

    for (let m = 0; m < n; m++) {
      // Build candidate pool (by index to guarantee uniqueness)
      let pool = questions.map((q, idx) => ({ q, idx }));

      if (opts.unique) {
        const remaining = pool.filter(({ idx }) => !globalUsedIndices.has(idx));
        if (remaining.length >= qCount) pool = remaining;
        // else: wrap around — reuse all questions
      }

      // Shuffle and pick exactly qCount — NO duplicates within a model (each idx appears once)
      const selected = shuffle(pool).slice(0, qCount);
      selected.forEach(({ idx }) => globalUsedIndices.add(idx));

      // Process each question
      const processedQ = selected.map(({ q }) => {
        const rawChoices = [
          { label: 'أ', val: (q[colMap.a] || '').trim() },
          { label: 'ب', val: (q[colMap.b] || '').trim() },
          { label: 'ج', val: (q[colMap.c] || '').trim() },
          { label: 'د', val: (q[colMap.d] || '').trim() }
        ];

        // Remove blank + duplicate choice texts
        const uniqueChoices = deduplicateChoices(rawChoices);

        // Shuffle choices
        const finalChoices = opts.shuffleA ? shuffle(uniqueChoices) : uniqueChoices;

        // Resolve correct answer after shuffling
        const correctRaw = (q[colMap.ans] || '').trim();
        const correctChoice = finalChoices.find(c =>
          c.val === correctRaw ||
          c.label === correctRaw ||
          c.label === mapEnToAr(correctRaw)
        );
        const correctLabel = correctChoice ? correctChoice.label : correctRaw;

        return {
          text:         (q[colMap.q] || '').trim(),
          choices:      finalChoices,
          correctLabel,
          topic:        (q[colMap.topic] || '').trim(),
          diff:         (q[colMap.diff]  || '').trim()
        };
      });

      models.push({
        name:      'نموذج ' + modelLetters[m],
        letter:    modelLetters[m],
        questions: opts.shuffleQ ? shuffle(processedQ) : processedQ
      });
    }

    btn.classList.remove('generating');
    renderModelTabs();
    goPanel(3);
    document.getElementById('gen-summary').textContent =
      `${n} نماذج × ${qCount} سؤال = ${n * qCount} سؤال إجمالي`;
    showToast(`تم توليد ${n} نماذج بنجاح`, 'success');
  }, 800);
}

// ══════════════════════════════════════════
// RENDER MODEL TABS
// ══════════════════════════════════════════
function renderModelTabs() {
  const tabs = document.getElementById('model-tabs');
  tabs.innerHTML =
    models.map((m, i) =>
      `<button class="exam-tab ${i===0?'active':''}" onclick="showModel(${i})">${m.name}</button>`
    ).join('')
    + `<button class="exam-tab" onclick="showAnswerKeys()">نماذج الإجابات</button>`;
  showModel(0);
}

function showModel(idx) {
  currentModelIdx = idx;
  document.querySelectorAll('.exam-tab').forEach((t,i) => t.classList.toggle('active', i===idx));
  document.getElementById('answer-key-view').style.display = 'none';
  document.getElementById('model-view').style.display      = 'block';
  document.getElementById('model-view').innerHTML          = renderExamPaper(models[idx], idx);
}

function showAnswerKeys() {
  document.querySelectorAll('.exam-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.exam-tab').item(models.length).classList.add('active');
  document.getElementById('model-view').style.display      = 'none';
  document.getElementById('answer-key-view').style.display = 'block';
  renderAllAnswerKeys();
}

// ══════════════════════════════════════════
// RENDER EXAM PAPER — in-app preview
// ══════════════════════════════════════════
function renderExamPaper(model, idx) {
  const fs  = design.fontSize;
  const hBg = design.headerBg;

  const header = opts.header ? `
    <div style="background:${hBg};color:${design.headerText};
                padding:1.5rem;text-align:center;margin:-2.5rem -2.5rem 1.5rem;">
      <div style="font-size:${fs+2}px;font-weight:700;margin-bottom:4px;">${examCfg.institution}</div>
      <div style="font-size:${fs+8}px;font-weight:900;font-family:'Amiri',serif;">امتحان مادة: ${examCfg.subject}</div>
      <div style="font-size:${fs-1}px;opacity:.85;margin-top:4px;">الزمن: ${examCfg.duration} دق | الدرجة: ${examCfg.grade}</div>
      ${opts.modelNum ? `<div style="margin-top:6px;display:inline-block;background:rgba(255,255,255,.25);
        padding:3px 16px;border-radius:20px;font-size:${fs+1}px;font-weight:700;">النموذج: ${model.letter}</div>` : ''}
    </div>` : '';

  const student = opts.student ? `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:1.5rem;font-size:${fs-1}px;">
      <div style="border-bottom:1px solid #2d3748;padding-bottom:4px;color:#4a5568;font-weight:500;">اسم الطالب: ________________</div>
      <div style="border-bottom:1px solid #2d3748;padding-bottom:4px;color:#4a5568;font-weight:500;">رقم الجلوس: ____________</div>
      <div style="border-bottom:1px solid #2d3748;padding-bottom:4px;color:#4a5568;font-weight:500;">الشعبة / الفصل: __________</div>
      <div style="border-bottom:1px solid #2d3748;padding-bottom:4px;color:#4a5568;font-weight:500;">التاريخ: ${examCfg.date||'___________'}</div>
    </div>` : '';

  const instructions = examCfg.instructions
    ? `<div style="background:#f8f9ff;border:1px solid #d0d8f0;border-radius:8px;
                   padding:10px 14px;margin-bottom:1.5rem;font-size:${fs-1}px;color:#4a5568;line-height:1.8;">
         <strong>التعليمات:</strong> ${examCfg.instructions}</div>` : '';

  const qItems = model.questions.map((item, i) => {
    const optHtml = item.choices.map(c => `
      <div style="display:flex;gap:6px;align-items:flex-start;font-size:${fs-1}px;color:#4a5568;margin-bottom:4px;">
        <span style="font-weight:700;color:#2d3748;flex-shrink:0;">${c.label})</span>
        <span>${c.val}</span>
      </div>`).join('');
    return `
      <div style="display:block;width:100%;margin:0 0 1.4rem 0;padding-bottom:1rem;
                  border-bottom:1px solid #e0e4ef;page-break-inside:avoid;">
        <div style="font-size:${fs}px;line-height:1.9;margin-bottom:8px;color:#1a1a2e;font-weight:500;">
          <strong style="color:${hBg};">${i+1}.</strong> ${item.text}
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px 16px;padding-right:16px;">${optHtml}</div>
      </div>`;
  }).join('');

  const answerTable = opts.answerTable ? buildAnswerFillTable(model.questions.length, hBg, fs) : '';

  return `
    <div class="exam-paper" id="paper-${idx}">
      ${header}${student}${instructions}
      <div style="display:block;width:100%;">${qItems}</div>
      ${answerTable}
    </div>`;
}

// ══════════════════════════════════════════
// ANSWER FILL TABLE — blank table at end of question paper
// ══════════════════════════════════════════
function buildAnswerFillTable(count, hBg, fs) {
  const cells = Array.from({ length: count }, (_, i) => `
    <td style="text-align:center;border:1px solid #c8d0e0;padding:8px 4px;min-width:44px;">
      <div style="font-size:9px;color:#888;margin-bottom:4px;">${i+1}</div>
      <div style="height:22px;border-bottom:1px solid #555;"></div>
    </td>`);
  const rows = [];
  for (let r = 0; r < cells.length; r += 10)
    rows.push(`<tr>${cells.slice(r, r+10).join('')}</tr>`);
  return `
    <div style="margin-top:2rem;page-break-inside:avoid;">
      <div style="background:${hBg};color:#fff;padding:6px 14px;font-size:${fs-1}px;
                  font-weight:700;border-radius:6px 6px 0 0;display:inline-block;">جدول الإجابات</div>
      <table style="width:100%;border-collapse:collapse;border:1px solid #c8d0e0;">${rows.join('')}</table>
    </div>`;
}

// ══════════════════════════════════════════
// RENDER ANSWER KEYS — in-app
// ══════════════════════════════════════════
function renderAllAnswerKeys() {
  document.getElementById('answer-keys-content').innerHTML = models.map(m => {
    const grid = m.questions.map((q, i) => `
      <div class="ans-item">
        <div class="ans-qnum">${i+1}</div>
        <div class="ans-val">${q.correctLabel}</div>
      </div>`).join('');
    return `
      <div style="margin-bottom:2rem;">
        <div style="font-size:15px;font-weight:700;margin-bottom:.75rem;color:var(--text);
                    display:flex;align-items:center;gap:8px;">
          <span style="background:${design.headerBg};color:${design.headerText};
                        padding:4px 14px;border-radius:20px;font-size:13px;">${m.name}</span>
          ${examCfg.subject} — ${m.questions.length} سؤال
        </div>
        <div style="display:grid;grid-template-columns:repeat(10,1fr);gap:8px;">${grid}</div>
      </div>`;
  }).join('');
}

// ══════════════════════════════════════════
// SHARED CSS FOR EXPORTED FILES
// ══════════════════════════════════════════
function exportPageCss(extra) {
  return `
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&family=Amiri:wght@400;700&display=swap');
    @page { size:A4; margin:18mm 20mm 20mm 20mm; }
    @page { @bottom-center {
      content: "صفحة " counter(page) " من " counter(pages);
      font-family:'Cairo',sans-serif; font-size:9pt; color:#666;
    }}
    * { box-sizing:border-box; margin:0; padding:0; }
    body { font-family:'Cairo',sans-serif; 