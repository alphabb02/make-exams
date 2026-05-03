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
  shuffleQ:       true,
  shuffleA:       true,
  unique:         true,
  header:         true,
  student:        true,
  modelNum:       true,
  answerKey:      true,
  answerTable:    true,
  answerTablePos: 'end'   // 'start' | 'end'
};

// Detect majority text direction from questions
function detectDir() {
  const sample = questions.slice(0, 30).map(q => (q[colMap.q]||'').trim()).join(' ');
  const arabic = (sample.match(/[\u0600-\u06FF]/g) || []).length;
  const latin  = (sample.match(/[A-Za-z]/g)         || []).length;
  return arabic >= latin ? 'rtl' : 'ltr';
}

let design = {
  headerBg:   '#1e3a8a',
  headerText: '#ffffff',
  fontSize:   14,
  layoutCols: 2
};

let examCfg = {};

// ══════════════════════════════════════════
// MATH & SYMBOL PROCESSING ENGINE
// Handles: LaTeX ($..$ $$..$$), Unicode math,
//          plain-text patterns (x^2, sqrt, fractions, etc.)
// ══════════════════════════════════════════

// ── Check if MathJax is available (loaded in index.html) ──
function hasMathJax() {
  return typeof window !== 'undefined' && window.MathJax && window.MathJax.typesetPromise;
}

// ── Detect if a string contains math content ──
function hasMath(text) {
  if (!text) return false;
  return (
    /\$[\s\S]+?\$/.test(text)          ||  // $...$ or $$...$$
    /\\\([\s\S]+?\\\)/.test(text)      ||  // \(...\)
    /\\\[[\s\S]+?\\\]/.test(text)      ||  // \[...\]
    /\\(?:frac|sqrt|sum|int|lim|vec|hat|bar|over|alpha|beta|gamma|delta|theta|pi|sigma|omega|lambda|mu|infty|cdot|times|div|pm|leq|geq|neq|approx|rightarrow|leftarrow|Rightarrow|forall|exists|in|subset|cup|cap|mathbb|mathbf|mathrm)\b/.test(text) ||
    /[²³¹⁰⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿ]/.test(text) ||  // Unicode superscripts
    /[₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎]/.test(text)   ||  // Unicode subscripts
    /[∑∏∫∂∇∞∈∉⊂⊃∪∩≤≥≠≈±×÷√∝∀∃∄⊕⊗∧∨¬→←↔⇒⇔αβγδεζηθικλμνξπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΠΡΣΤΥΦΧΨΩ]/.test(text) ||
    /\b(?:sqrt|log|ln|sin|cos|tan|cot|sec|csc|lim|max|min)\s*[\(\d]/.test(text) ||
    /\d+\s*[\^]\s*\d+/.test(text)      ||  // x^2
    /\d+\s*\/\s*\d+/.test(text)            // fractions like 3/4
  );
}

// ── Convert plain-text math patterns to LaTeX ──
function normalizeToLatex(text) {
  if (!text) return text;
  let t = text;

  // Already has LaTeX delimiters — leave as-is
  if (/\$|\\[(\[]/.test(t)) return t;

  // Unicode superscripts → LaTeX ^{}
  const supMap = {'²':'^{2}','³':'^{3}','¹':'^{1}','⁰':'^{0}','⁴':'^{4}','⁵':'^{5}','⁶':'^{6}','⁷':'^{7}','⁸':'^{8}','⁹':'^{9}','ⁿ':'^{n}','⁺':'^{+}','⁻':'^{-}'};
  const subMap = {'₀':'_{0}','₁':'_{1}','₂':'_{2}','₃':'_{3}','₄':'_{4}','₅':'_{5}','₆':'_{6}','₇':'_{7}','₈':'_{8}','₉':'_{9}'};

  let hasSup = false, hasSub = false;
  for (const [k,v] of Object.entries(supMap)) { if (t.includes(k)) { t = t.replaceAll(k,v); hasSup=true; } }
  for (const [k,v] of Object.entries(subMap)) { if (t.includes(k)) { t = t.replaceAll(k,v); hasSub=true; } }

  // x^2 plain text → keep as LaTeX inline
  if (/[a-zA-Z0-9]\^[{0-9a-zA-Z]/.test(t)) hasSup = true;

  // sqrt(...) → \sqrt{...}
  t = t.replace(/\bsqrt\s*\(([^)]+)\)/g, '\\sqrt{$1}');
  t = t.replace(/\bsqrt\s+(\S+)/g, '\\sqrt{$1}');
  t = t.replace(/√\s*\(([^)]+)\)/g, '\\sqrt{$1}');
  t = t.replace(/√\s*(\S+)/g, '\\sqrt{$1}');

  // n/d fractions in math context: wrap whole expression
  // Only wrap if looks like standalone math (not inside prose like "1/3 of students")
  t = t.replace(/\b(\d+)\s*\/\s*(\d+)\b/g, '\\frac{$1}{$2}');

  // Greek letters already present as Unicode → wrap in $
  const needsWrap = hasSup || hasSub ||
    /\\(?:sqrt|frac|sum|int|alpha|beta|gamma|delta|theta|pi|sigma|omega|lambda|sin|cos|tan|log|ln|lim)/.test(t) ||
    /[∑∏∫∂∇∞∈∉⊂⊃∪∩≤≥≠≈±×÷√∝⊕⊗∧∨→←↔⇒⇔]/.test(t);

  if (needsWrap && !/\$/.test(t)) {
    // Wrap only the math parts (sequences of math tokens) in $...$
    // Strategy: wrap entire string if it looks fully mathematical
    const wordCount = t.trim().split(/\s+/).length;
    const mathTokens = (t.match(/[\\{}\^_]|\\[a-zA-Z]+/g)||[]).length;
    if (mathTokens > 0 && (mathTokens / wordCount > 0.3 || wordCount <= 5)) {
      t = `$${t}$`;
    } else {
      // Wrap just the mathematical sub-expressions
      t = t.replace(/((?:[a-zA-Z0-9][\^_][{}\d\w]*|\\[a-zA-Z]+(?:\{[^}]*\})*)+)/g, '$$$1$$$');
    }
  }

  return t;
}

// ── Full render pipeline: normalize + sanitize HTML + keep structure ──
function renderMath(text) {
  if (!text) return '';
  // Preserve existing HTML tags (bold, etc.)
  const hasHtml = /<[a-zA-Z]/.test(text);
  if (hasHtml) {
    // Process text nodes only, leave tags intact
    return text.replace(/(?<=>|^)([^<]+)(?=<|$)/g, (match) => renderMath(match));
  }
  if (!hasMath(text)) return text;
  return normalizeToLatex(text);
}

// ── After inserting HTML, tell MathJax to typeset ──
function typesetMath(element) {
  if (!hasMathJax()) return;
  MathJax.typesetPromise([element]).catch(err => console.warn('MathJax:', err));
}

// ── Build MathJax-enabled HTML for exported files ──
function mathJaxScript() {
  return `
  <script>
    window.MathJax = {
      tex: {
        inlineMath: [['$','$'], ['\\\\(','\\\\)']],
        displayMath: [['$$','$$'], ['\\\\[','\\\\]']],
        packages: {'[+]': ['ams','boldsymbol']},
        tags: 'none'
      },
      options: { skipHtmlTags: ['script','noscript','style','textarea','pre'] },
      startup: { typeset: true }
    };
  <\/script>
  <script src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js" id="MathJax-script" async><\/script>`;
}

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

function setAnswerTablePos(val) {
  opts.answerTablePos = val;
  const btnEnd   = document.getElementById('pos-end');
  const btnStart = document.getElementById('pos-start');
  if (!btnEnd || !btnStart) return;
  if (val === 'end') {
    btnEnd.style.background   = 'var(--accent)'; btnEnd.style.color   = '#fff'; btnEnd.style.borderColor   = 'var(--accent)';
    btnStart.style.background = 'transparent';   btnStart.style.color = 'var(--text2)'; btnStart.style.borderColor = 'var(--border)';
  } else {
    btnStart.style.background = 'var(--accent)'; btnStart.style.color = '#fff'; btnStart.style.borderColor = 'var(--accent)';
    btnEnd.style.background   = 'transparent';   btnEnd.style.color   = 'var(--text2)'; btnEnd.style.borderColor   = 'var(--border)';
  }
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
  // Now labels are A/B/C/D — map Arabic or lowercase to uppercase
  return { A:'A', B:'B', C:'C', D:'D', a:'A', b:'B', c:'C', d:'D',
           'أ':'A', 'ب':'B', 'ج':'C', 'د':'D' }[letter] || letter.toUpperCase() || letter;
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
          { label: 'A', val: (q[colMap.a] || '').trim() },
          { label: 'B', val: (q[colMap.b] || '').trim() },
          { label: 'C', val: (q[colMap.c] || '').trim() },
          { label: 'D', val: (q[colMap.d] || '').trim() }
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
        <span style="font-weight:700;color:#2d3748;flex-shrink:0;min-width:18px;">${c.label})</span>
        <span>${renderMath(c.val)}</span>
      </div>`).join('');
    return `
      <div style="display:block;width:100%;margin:0 0 1.4rem 0;padding-bottom:1rem;
                  border-bottom:1px solid #e0e4ef;page-break-inside:avoid;">
        <div style="font-size:${fs}px;line-height:1.9;margin-bottom:8px;color:#1a1a2e;font-weight:500;">
          <span style="font-weight:700;">${i+1})</span> ${renderMath(item.text)}
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px 16px;padding-right:24px;">${optHtml}</div>
      </div>`;
  }).join('');

  const answerTable = opts.answerTable ? buildAnswerFillTable(model.questions.length, hBg, fs) : '';
  const dir = detectDir();
  const tableAtStart = opts.answerTablePos === 'start';

  const paperHtml = `
    <div class="exam-paper" id="paper-${idx}" style="direction:${dir};">
      ${header}${student}${instructions}
      ${tableAtStart ? answerTable : ''}
      <div style="display:block;width:100%;">${qItems}</div>
      ${tableAtStart ? '' : answerTable}
    </div>`;

  // Typeset math after DOM insertion (called by showModel)
  setTimeout(() => {
    const el = document.getElementById('paper-' + idx);
    if (el) typesetMath(el);
  }, 50);

  return paperHtml;
}

// ══════════════════════════════════════════
// ANSWER FILL TABLE — blank table at end of question paper
// ══════════════════════════════════════════
function buildAnswerFillTable(count, hBg, fs) {
  // Always LTR, left-to-right numbering
  // Two rows per group: "Question" row (numbers) + "Answer" row (blank)
  const perRow = 10;
  let tableRows = '';
  for (let r = 0; r < count; r += perRow) {
    const nums = [];
    const blanks = [];
    for (let c = r; c < Math.min(r + perRow, count); c++) {
      nums.push(`<td style="text-align:center;border:1px solid #c8d0e0;padding:4px 2px;
                             min-width:40px;font-size:8.5pt;font-weight:700;color:#333;">${c+1}</td>`);
      blanks.push(`<td style="border:1px solid #c8d0e0;padding:0;min-width:40px;height:22px;"></td>`);
    }
    // label column on left
    tableRows += `
      <tr>
        <td style="border:1px solid #c8d0e0;padding:4px 6px;font-size:8.5pt;font-weight:700;
                   color:#fff;background:${hBg};white-space:nowrap;width:60px;">Question</td>
        ${nums.join('')}
      </tr>
      <tr>
        <td style="border:1px solid #c8d0e0;padding:4px 6px;font-size:8.5pt;font-weight:700;
                   color:#fff;background:${hBg};white-space:nowrap;">Answer</td>
        ${blanks.join('')}
      </tr>`;
  }
  return `
    <div style="margin-top:1.5rem;page-break-inside:avoid;direction:ltr;">
      <table style="border-collapse:collapse;border:1px solid #c8d0e0;">${tableRows}</table>
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
    body { font-family:'Cairo',sans-serif; background:#fff; color:#1a1a2e; direction:rtl; font-size:13px; }
    @media print { body { -webkit-print-color-adjust:exact; print-color-adjust:exact; } }
    .q-block { page-break-inside: avoid; }
    ${extra||''}`;
}

// ══════════════════════════════════════════
// BUILD QUESTION PAPER HTML (A4)
// ══════════════════════════════════════════
function buildQuestionPaperHtml(model) {
  const hBg = design.headerBg, hTx = design.headerText;

  const dir = detectDir();
  const qHtml = model.questions.map((item, i) => {
    const pairs = [];
    for (let r = 0; r < item.choices.length; r += 2) {
      const c1 = item.choices[r], c2 = item.choices[r+1];
      pairs.push(`<tr>
        <td style="width:50%;padding:2px 8px;font-size:9.5pt;color:#333;vertical-align:top;">
          ${c1 ? `<b>${c1.label})</b> ${renderMath(c1.val)}` : ''}
        </td>
        <td style="width:50%;padding:2px 8px;font-size:9.5pt;color:#333;vertical-align:top;">
          ${c2 ? `<b>${c2.label})</b> ${renderMath(c2.val)}` : ''}
        </td></tr>`);
    }
    return `<div class="q-block" style="margin-bottom:9px;padding-bottom:8px;border-bottom:1px solid #dde2ee;">
      <div style="font-size:10.5pt;line-height:1.75;color:#1a1a2e;margin-bottom:4px;">
        <b>${i+1})</b> ${renderMath(item.text)}
      </div>
      <table style="width:100%;border-collapse:collapse;">${pairs.join('')}</table>
    </div>`;
  }).join('');

  // Fill table always LTR, position controlled by opt
  let fillTableHtml = '';
  if (opts.answerTable) {
    const perRow = 10;
    let tRows = '';
    for (let r = 0; r < model.questions.length; r += perRow) {
      const nums = [], blanks = [];
      for (let c = r; c < Math.min(r+perRow, model.questions.length); c++) {
        nums.push(`<td style="text-align:center;border:1px solid #c8d0e0;padding:4px 2px;
                               min-width:38px;font-size:8.5pt;font-weight:700;color:#333;">${c+1}</td>`);
        blanks.push(`<td style="border:1px solid #c8d0e0;min-width:38px;height:20px;"></td>`);
      }
      tRows += `<tr>
        <td style="border:1px solid #c8d0e0;padding:3px 6px;font-size:8.5pt;font-weight:700;
                   color:#fff;background:${hBg};white-space:nowrap;">Question</td>${nums.join('')}
      </tr><tr>
        <td style="border:1px solid #c8d0e0;padding:3px 6px;font-size:8.5pt;font-weight:700;
                   color:#fff;background:${hBg};white-space:nowrap;">Answer</td>${blanks.join('')}
      </tr>`;
    }
    fillTableHtml = `<div style="margin-top:14pt;page-break-inside:avoid;direction:ltr;">
      <table style="border-collapse:collapse;border:1px solid #c8d0e0;">${tRows}</table>
    </div>`;
  }

  const tableAtStart = opts.answerTablePos === 'start';

  return `<!DOCTYPE html><html lang="${dir==='ltr'?'en':'ar'}" dir="${dir}"><head><meta charset="UTF-8">
<style>${exportPageCss()}</style>
${mathJaxScript()}
</head><body style="direction:${dir};">
  <div style="background:${hBg};color:${hTx};padding:14px 20px;text-align:center;margin-bottom:12px;">
    <div style="font-size:13pt;font-weight:700;">${examCfg.institution||''}</div>
    <div style="font-size:17pt;font-weight:900;font-family:'Amiri',serif;margin:4px 0;">امتحان مادة: ${examCfg.subject||''}</div>
    <div style="font-size:10pt;opacity:.9;">الزمن: ${examCfg.duration} دقيقة | الدرجة: ${examCfg.grade} | ${examCfg.date||''}</div>
    <div style="margin-top:5px;display:inline-block;background:rgba(255,255,255,.25);
                padding:2px 18px;border-radius:20px;font-size:11pt;font-weight:700;">النموذج: ${model.letter}</div>
  </div>
  <table style="width:100%;border-collapse:collapse;font-size:10pt;margin-bottom:10px;"><tr>
    <td style="padding:4px 8px;border-bottom:1px solid #333;width:50%;">اسم الطالب: _______________________</td>
    <td style="padding:4px 8px;border-bottom:1px solid #333;">رقم الجلوس: ________________</td>
  </tr><tr>
    <td style="padding:4px 8px;border-bottom:1px solid #333;">الشعبة / الفصل: __________________</td>
    <td style="padding:4px 8px;border-bottom:1px solid #333;">التاريخ: ${examCfg.date||'_______________'}</td>
  </tr></table>
  ${examCfg.instructions ? `<div style="background:#f5f7ff;border:1px solid #c8d0f0;border-radius:5px;
    padding:7px 12px;margin-bottom:10px;font-size:9.5pt;color:#333;line-height:1.7;">
    <strong>التعليمات:</strong> ${examCfg.instructions}</div>` : ''}
  ${tableAtStart ? fillTableHtml : ''}
  ${qHtml}
  ${tableAtStart ? '' : fillTableHtml}
</body></html>`;
}

// ══════════════════════════════════════════
// BUILD ANSWER SHEET HTML (A4)
// Correct option gets green highlight (marker)
// ══════════════════════════════════════════
function buildAnswerSheetHtml(model) {
  const hBg = design.headerBg, hTx = design.headerText;

  const extraCss = `
    .opt-correct {
      background: #bbf7d0;
      border: 1.5px solid #16a34a;
      border-radius: 4px;
      padding: 1px 6px;
      display: inline-flex;
      gap: 5px;
      align-items: flex-start;
    }
    .opt-correct b { color: #15803d; }
  `;

  const qHtml = model.questions.map((item, i) => {
    const opts2 = item.choices.map(c => {
      const isCorrect = c.label === item.correctLabel;
      return `<td style="width:50%;padding:2px 8px;font-size:9.5pt;vertical-align:top;">
        <span class="${isCorrect ? 'opt-correct' : ''}"
              style="display:inline-flex;gap:5px;align-items:flex-start;color:#333;">
          <b style="${isCorrect ? 'color:#15803d;' : ''};flex-shrink:0;">${c.label})</b>${renderMath(c.val)}
        </span></td>`;
    });
    const rows = [];
    for (let r = 0; r < opts2.length; r += 2)
      rows.push(`<tr>${opts2[r]||'<td></td>'}${opts2[r+1]||'<td></td>'}</tr>`);
    return `<div class="q-block" style="margin-bottom:9px;padding-bottom:8px;border-bottom:1px solid #dde2ee;">
      <div style="font-size:10.5pt;line-height:1.75;color:#1a1a2e;margin-bottom:4px;">
        <b>${i+1})</b> ${renderMath(item.text)}
      </div>
      <table style="width:100%;border-collapse:collapse;margin-right:16px;">${rows.join('')}</table>
    </div>`;
  }).join('');

  // Summary grid
  const keyCells = model.questions.map((q, i) => `
    <td style="text-align:center;border:1px solid #bbf7d0;padding:5px 3px;min-width:40px;background:#f0fdf4;">
      <div style="font-size:8pt;color:#666;">${i+1}</div>
      <div style="font-size:12pt;font-weight:700;color:#15803d;">${q.correctLabel}</div>
    </td>`);
  const keyRows = [];
  for (let r = 0; r < keyCells.length; r += 10)
    keyRows.push(`<tr>${keyCells.slice(r, r+10).join('')}</tr>`);

  const dir = detectDir();
  return `<!DOCTYPE html><html lang="${dir==='ltr'?'en':'ar'}" dir="${dir}"><head><meta charset="UTF-8">
<style>${exportPageCss(extraCss)}</style>
${mathJaxScript()}
</head><body style="direction:${dir};">
  <div style="background:${hBg};color:${hTx};padding:14px 20px;text-align:center;margin-bottom:12px;
              outline:3px solid #16a34a;outline-offset:-3px;">
    <div style="font-size:13pt;font-weight:700;">${examCfg.institution||''}</div>
    <div style="font-size:17pt;font-weight:900;font-family:'Amiri',serif;margin:4px 0;">امتحان مادة: ${examCfg.subject||''}</div>
    <div style="font-size:10pt;opacity:.9;">الزمن: ${examCfg.duration} دقيقة | الدرجة: ${examCfg.grade} | ${examCfg.date||''}</div>
    <div style="margin-top:5px;display:inline-flex;gap:8px;align-items:center;justify-content:center;flex-wrap:wrap;">
      <span style="display:inline-block;background:rgba(255,255,255,.25);padding:2px 18px;
                   border-radius:20px;font-size:11pt;font-weight:700;">النموذج: ${model.letter}</span>
      <span style="display:inline-block;background:#16a34a;color:#fff;padding:2px 18px;
                   border-radius:20px;font-size:10pt;font-weight:700;">نموذج الإجابات</span>
    </div>
  </div>
  ${qHtml}
  <div style="margin-top:18pt;page-break-inside:avoid;">
    <div style="background:#16a34a;color:#fff;padding:5pt 12pt;font-size:10pt;font-weight:700;
                border-radius:4pt 4pt 0 0;display:inline-block;">ملخص الإجابات الصحيحة</div>
    <table style="width:100%;border-collapse:collapse;border:1px solid #bbf7d0;">${keyRows.join('')}</table>
  </div>
</body></html>`;
}

// ══════════════════════════════════════════
// EXPORT — PDF ZIP
// ══════════════════════════════════════════
async function exportAllPDF() {
  if (!models.length) { showToast('لا توجد نماذج بعد', 'error'); return; }
  showToast('جاري إنشاء ملفات PDF...', '');

  const zip       = new JSZip();
  const qFolder   = zip.folder('الأسئلة');
  const ansFolder = zip.folder('نماذج_الإجابات');

  models.forEach(m => {
    qFolder.file(`${m.name} - أسئلة.html`,    buildQuestionPaperHtml(m));
    ansFolder.file(`${m.name} - إجابات.html`, buildAnswerSheetHtml(m));
  });

  // Combined questions file
  let combined = `<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="UTF-8">
<style>
  ${exportPageCss()}
  .model-wrap { page-break-after: always; }
</style></head><body>`;
  models.forEach(m => {
    const inner = buildQuestionPaperHtml(m)
      .replace(/[\s\S]*?<body[^>]*>/i,'').replace(/<\/body>[\s\S]*/i,'');
    combined += `<div class="model-wrap">${inner}</div>`;
  });
  combined += '</body></html>';
  zip.file('كل_النماذج_مجمعة.html', combined);

  const blob = await zip.generateAsync({ type: 'blob' });
  triggerDownload(blob, `نماذج_PDF_${examCfg.subject||'امتحان'}.zip`);
  showToast('ZIP جاهز — افتح HTML واضغط Ctrl+P للطباعة كـ PDF', 'success');
}

// ══════════════════════════════════════════
// EXPORT — WORD ZIP
// ══════════════════════════════════════════
async function exportAllWord() {
  if (!models.length) { showToast('لا توجد نماذج بعد', 'error'); return; }
  showToast('جاري إنشاء ملفات Word...', '');

  const zip       = new JSZip();
  const qFolder   = zip.folder('الأسئلة');
  const ansFolder = zip.folder('نماذج_الإجابات');
  const hBg = design.headerBg, hTx = design.headerText;

  const dir = detectDir();
  const wordWrap = (body, modelLetter, isAnswer) =>
    `<html xmlns:o="urn:schemas-microsoft-com:office:office"
           xmlns:w="urn:schemas-microsoft-com:office:word"
           xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8"><style>
  @page Section1 { size:21cm 29.7cm; margin:18mm 20mm 20mm 20mm;
    mso-header-margin:10mm; mso-footer-margin:10mm; mso-page-numbers:1; }
  div.Section1 { page:Section1; }
  body { font-family:'Arial Unicode MS',Arial,sans-serif; direction:${dir}; font-size:10pt; }
  p { margin:0 0 3pt; text-align:${dir==='ltr'?'left':'right'}; }
  td { text-align:${dir==='ltr'?'left':'right'}; }
  .correct { background:#bbf7d0; border:1pt solid #16a34a; padding:1pt 4pt; }
</style></head>
<body dir="${dir}"><div class="Section1">
  <div style="background:${hBg};color:${hTx};padding:10pt;text-align:center;margin-bottom:10pt;
              ${isAnswer ? 'border:3pt solid #16a34a;' : ''}">
    <div style="font-size:13pt;font-weight:700;">${examCfg.institution||''}</div>
    <div style="font-size:17pt;font-weight:900;">${examCfg.subject||''}</div>
    <div style="font-size:10pt;">Duration: ${examCfg.duration} min &nbsp;|&nbsp; Grade: ${examCfg.grade} &nbsp;|&nbsp; ${examCfg.date||''}</div>
    <div style="font-size:11pt;font-weight:700;">Model: ${modelLetter}${isAnswer ? ' — Answer Key' : ''}</div>
  </div>
  ${body}
</div></body></html>`;

  models.forEach(m => {
    // ── Question paper body ──
    const isLtr = dir === 'ltr';
    const studentTable = `<table width="100%" style="margin-bottom:8pt;" dir="${dir}"><tr>
      <td width="50%" style="border-bottom:1pt solid #333;padding:3pt;font-size:10pt;">${isLtr ? 'Student Name' : 'اسم الطالب'}: _______________________</td>
      <td style="border-bottom:1pt solid #333;padding:3pt;font-size:10pt;">${isLtr ? 'Seat No.' : 'رقم الجلوس'}: ________________</td>
    </tr><tr>
      <td style="border-bottom:1pt solid #333;padding:3pt;font-size:10pt;">${isLtr ? 'Class / Section' : 'الشعبة'}: __________________</td>
      <td style="border-bottom:1pt solid #333;padding:3pt;font-size:10pt;">${isLtr ? 'Date' : 'التاريخ'}: ${examCfg.date||'___'}</td>
    </tr></table>`;

    const instrBlock = examCfg.instructions
      ? `<p style="background:#f5f7ff;border:1pt solid #c0c8e0;padding:5pt;font-size:9.5pt;margin-bottom:8pt;">
           <b>التعليمات:</b> ${examCfg.instructions}</p>` : '';

    const qHtml = m.questions.map((item, i) => {
      const pairs = [];
      for (let r = 0; r < item.choices.length; r += 2) {
        const c1 = item.choices[r], c2 = item.choices[r+1];
        pairs.push(`<tr>
          <td width="50%" style="padding:2pt 6pt;font-size:10pt;">${c1 ? `<b>${c1.label})</b> ${c1.val}` : ''}</td>
          <td width="50%" style="padding:2pt 6pt;font-size:10pt;">${c2 ? `<b>${c2.label})</b> ${c2.val}` : ''}</td>
        </tr>`);
      }
      return `<div dir="${dir}" style="margin-bottom:6pt;">
        <p style="margin:0 0 4pt;font-size:10.5pt;line-height:1.75;"><b>${i+1})</b> ${item.text}</p>
        <table width="100%" dir="${dir}" style="border-collapse:collapse;margin-bottom:6pt;">${pairs.join('')}</table>
        <hr style="border:none;border-top:1px solid #dde;margin:0 0 4pt;">
      </div>`;
    }).join('');

    let fillTable = '';
    if (opts.answerTable) {
      const perRow = 10;
      let tRows = '';
      for (let r = 0; r < m.questions.length; r += perRow) {
        const nums = [], blanks = [];
        for (let c = r; c < Math.min(r+perRow, m.questions.length); c++) {
          nums.push(`<td style="text-align:center;border:1pt solid #c8d0e0;padding:3pt;width:36pt;font-size:9pt;font-weight:700;">${c+1}</td>`);
          blanks.push(`<td style="border:1pt solid #c8d0e0;width:36pt;height:16pt;"></td>`);
        }
        tRows += `<tr>
          <td style="border:1pt solid #c8d0e0;padding:3pt 6pt;font-size:9pt;font-weight:700;
                     background:${hBg};color:#fff;white-space:nowrap;">Question</td>${nums.join('')}
        </tr><tr>
          <td style="border:1pt solid #c8d0e0;padding:3pt 6pt;font-size:9pt;font-weight:700;
                     background:${hBg};color:#fff;white-space:nowrap;">Answer</td>${blanks.join('')}
        </tr>`;
      }
      const tablePos = opts.answerTablePos === 'start' ? 'start' : 'end';
      fillTable = `<br><table dir="ltr" style="border-collapse:collapse;direction:ltr;">${tRows}</table>`;
      if (tablePos === 'start') fillTable = fillTable + '<!--FILLTABLE_START-->';
    }

    const wordBody = fillTable.includes('<!--FILLTABLE_START-->')
      ? fillTable.replace('<!--FILLTABLE_START-->','') + studentTable + instrBlock + qHtml
      : studentTable + instrBlock + qHtml + fillTable;
    qFolder.file(`${m.name} - أسئلة.doc`,
      '\ufeff' + wordWrap(wordBody, m.letter, false));

    // ── Answer sheet body ──
    const aHtml = m.questions.map((item, i) => {
      const pairs = [];
      for (let r = 0; r < item.choices.length; r += 2) {
        const c1 = item.choices[r], c2 = item.choices[r+1];
        const cell = c => {
          if (!c) return '<td width="50%"></td>';
          const ok = c.label === item.correctLabel;
          return `<td width="50%" style="padding:2pt 6pt;font-size:10pt;">
            <span${ok ? ' class="correct" style="background:#bbf7d0;padding:1pt 4pt;"' : ''}>
              <b${ok ? ' style="color:#15803d;"' : ''}>${c.label})</b> ${c.val}
            </span></td>`;
        };
        pairs.push(`<tr>${cell(c1)}${cell(c2)}</tr>`);
      }
      return `<p style="margin:0 0 4pt;font-size:10.5pt;line-height:1.75;"><b>${i+1})</b> ${item.text}</p>
        <table width="100%" style="border-collapse:collapse;margin-bottom:6pt;">${pairs.join('')}</table>
        <hr style="border:none;border-top:1px solid #dde;margin:0 0 5pt;">`;
    }).join('');

    const keyCells = m.questions.map((q,i) =>
      `<td style="text-align:center;border:1pt solid #bbf7d0;padding:4pt;width:40pt;background:#f0fdf4;">
        <div style="font-size:8pt;color:#666;">${i+1}</div>
        <div style="font-size:12pt;font-weight:700;color:#15803d;">${q.correctLabel}</div>
      </td>`);
    const keyRows = [];
    for (let r=0; r < keyCells.length; r+=10) keyRows.push(`<tr>${keyCells.slice(r,r+10).join('')}</tr>`);
    const keyGrid = `<br><div style="background:#16a34a;color:#fff;padding:4pt 10pt;font-size:10pt;font-weight:700;">ملخص الإجابات</div>
      <table style="border-collapse:collapse;width:100%;">${keyRows.join('')}</table>`;

    ansFolder.file(`${m.name} - إجابات.doc`,
      '\ufeff' + wordWrap(aHtml + keyGrid, m.letter, true));
  });

  const blob = await zip.generateAsync({ type: 'blob' });
  triggerDownload(blob, `نماذج_Word_${examCfg.subject||'امتحان'}.zip`);
  showToast('ZIP جاهز — كل نموذج في ملفين (أسئلة + إجابات)', 'success');
}

// ══════════════════════════════════════════
// EXPORT — ANSWER KEYS ONLY ZIP
// ══════════════════════════════════════════
async function exportAnswerKeys() {
  if (!models.length) { showToast('لا توجد نماذج بعد', 'error'); return; }
  const zip    = new JSZip();
  const folder = zip.folder('نماذج_الإجابات');
  models.forEach(m => folder.file(`${m.name} - إجابات.html`, buildAnswerSheetHtml(m)));
  const blob = await zip.generateAsync({ type: 'blob' });
  triggerDownload(blob, `نماذج_الإجابات_${examCfg.subject||'امتحان'}.zip`);
  showToast('تم تنزيل نماذج الإجابات', 'success');
}

// ══════════════════════════════════════════
// UTILITIES
// ══════════════════════════════════════════
function triggerDownload(blob, filename) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}

function showToast(msg, type = '') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className   = 'toast show ' + type;
  setTimeout(() => t.classList.remove('show'), 3500);
}

// ══════════════════════════════════════════
// INIT
// ══════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  initDropZone();
  initColorPickers();
});
