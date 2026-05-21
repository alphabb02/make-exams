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
  answerTablePos: 'start'  // always before questions by default
};

// Detect majority text direction from questions
function detectDir() {
  const sample = questions.slice(0, 30).map(q => (q[colMap.q]||'').trim()).join(' ');
  const arabic = (sample.match(/[\u0600-\u06FF]/g) || []).length;
  const latin  = (sample.match(/[A-Za-z]/g)         || []).length;
  return arabic >= latin ? 'rtl' : 'ltr';
}

let design = {
  headerBg:     '#1e3a8a',
  headerText:   '#ffffff',
  fontSize:     14,
  layoutCols:   2,
  answerStyle:  'normal'  // 'normal' or 'highlighted'
};

let examCfg = {};

// ══════════════════════════════════════════
// PROJECT STORAGE SYSTEM
// ══════════════════════════════════════════
const STORAGE_KEY = 'exam_projects_v1';
const ENCRYPTION_KEY = 'exam_builder_2026';

// Simple encryption using base64 + XOR
function encryptProject(data) {
  try {
    const json = JSON.stringify(data);
    const base64 = btoa(unescape(encodeURIComponent(json)));
    let encrypted = '';
    for (let i = 0; i < base64.length; i++) {
      encrypted += String.fromCharCode(base64.charCodeAt(i) ^ ENCRYPTION_KEY.charCodeAt(i % ENCRYPTION_KEY.length));
    }
    return btoa(encrypted);
  } catch (e) {
    console.error('Encryption error:', e);
    return null;
  }
}

function decryptProject(encrypted) {
  try {
    const base64 = atob(encrypted);
    let decrypted = '';
    for (let i = 0; i < base64.length; i++) {
      decrypted += String.fromCharCode(base64.charCodeAt(i) ^ ENCRYPTION_KEY.charCodeAt(i % ENCRYPTION_KEY.length));
    }
    const json = decodeURIComponent(escape(atob(decrypted)));
    return JSON.parse(json);
  } catch (e) {
    console.error('Decryption error:', e);
    return null;
  }
}

// Get all saved projects
function getAllProjects() {
  try {
    const data = localStorage.getItem(STORAGE_KEY);
    return data ? JSON.parse(data) : [];
  } catch (e) {
    console.error('Get projects error:', e);
    return [];
  }
}

// Save project
function saveProjectToStorage(projectData) {
  try {
    const projects = getAllProjects();
    const encrypted = encryptProject(projectData);
    if (!encrypted) {
      console.warn('Encryption failed');
      return null;
    }

    // Stable ID: reuse existing if available, else generate new
    const projectId = projectData.id || Date.now();
    const subjectName = projectData.examCfg?.subject || 'مشروع بدون اسم';

    const project = {
      id: projectId,
      name: subjectName,
      questions: projectData.questions?.length || 0,
      models: projectData.models?.length || 0,
      date: new Date().toISOString(),
      encrypted: encrypted
    };

    // Find by exact ID to update in place
    const existingIdx = projects.findIndex(p => p.id === projectId);
    if (existingIdx >= 0) {
      projects[existingIdx] = project;
    } else {
      projects.push(project);
    }

    const storageData = JSON.stringify(projects);
    const sizeInKB = new Blob([storageData]).size / 1024;
    if (sizeInKB > 5000) {
      console.warn(`Storage size: ${sizeInKB.toFixed(0)}KB - approaching limit`);
    }

    localStorage.setItem(STORAGE_KEY, storageData);
    console.log(`Project saved: "${subjectName}" — ${project.questions} questions, ${project.models} models`);
    return project.id;
  } catch (e) {
    console.error('Save error:', e);
    return null;
  }
}

// Load project from storage
function loadProjectFromStorage(id) {
  try {
    const projects = getAllProjects();
    const project = projects.find(p => p.id === id);
    if (project && project.encrypted) {
      return decryptProject(project.encrypted);
    }
    return null;
  } catch (e) {
    console.error('Load error:', e);
    return null;
  }
}

// Delete project
function deleteProjectFromStorage(id) {
  try {
    let projects = getAllProjects();
    projects = projects.filter(p => p.id !== id);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(projects));
    return true;
  } catch (e) {
    console.error('Delete error:', e);
    return false;
  }
}

// Auto-save
function autoSaveCurrentProject() {
  if (questions.length === 0) return;

  // Generate a stable ID for this session's project if not yet assigned
  if (!window.currentProjectId) {
    window.currentProjectId = Date.now();
  }

  // Read current subject name from form (if filled) so it is always up-to-date
  const liveSubject = document.getElementById('subjectName')?.value?.trim() || examCfg.subject || '';
  if (liveSubject) examCfg.subject = liveSubject;

  const projectData = {
    id: window.currentProjectId,
    questions: questions,
    colMap: colMap,
    opts: opts,
    design: design,
    examCfg: examCfg,
    models: models
  };
  window.currentProjectId = saveProjectToStorage(projectData);
}

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
// MULTI-FILE QUESTION BANK SYSTEM
// Allow importing multiple Excel files and distributing questions across models
// ══════════════════════════════════════════

let bankFiles = [];  // [{name, questions, alloc}]
let multiFileMode = false;

function toggleMultiFileMode(enabled) {
  multiFileMode = enabled;
  const container = document.getElementById('multi-file-container');
  if (!container) return;
  container.style.display = enabled ? 'block' : 'none';
  
  // Hide/show single file uploader
  const singleUpload = document.getElementById('dropZone');
  if (singleUpload) singleUpload.style.display = enabled ? 'none' : 'block';
}

function addFileToBank(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
      
      if (raw.length < 2) { showToast('الملف فارغ', 'error'); return; }
      
      const headers = raw[0].map(h => String(h || '').trim());
      const fileQuestions = raw.slice(1)
        .filter(r => r.some(c => c !== undefined && c !== ''))
        .map(r => {
          const obj = {};
          headers.forEach((h, i) => obj[h] = r[i] !== undefined ? String(r[i]).trim() : '');
          return obj;
        });
      
      bankFiles.push({
        name: file.name.replace(/\.[^/.]+$/, ''),
        questions: fileQuestions,
        alloc: fileQuestions.length,
        headers
      });
      
      renderBankFilesList();
      showToast(`✓ تم إضافة "${file.name}" (${fileQuestions.length} سؤال)`, 'success');
    } catch(err) {
      showToast('خطأ في قراءة الملف: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function removeBankFile(idx) {
  bankFiles.splice(idx, 1);
  renderBankFilesList();
}

function updateBankFileAlloc(idx, val) {
  const num = parseInt(val) || 0;
  if (bankFiles[idx]) bankFiles[idx].alloc = Math.min(num, bankFiles[idx].questions.length);
  renderBankFilesList();
}

function renderBankFilesList() {
  const container = document.getElementById('bank-files-list');
  if (!container) return;
  
  const total = bankFiles.reduce((s, f) => s + f.alloc, 0);
  const qPerModel = parseInt(document.getElementById('qPerModel')?.value || 20);
  const status = total >= qPerModel ? '✓' : '⚠️';
  const color = total >= qPerModel ? 'var(--accent3)' : 'var(--warn)';
  
  container.innerHTML = `
    <div style="margin-bottom:12px;padding:10px;background:var(--bg3);border-radius:8px;">
      <div style="font-size:13px;color:var(--text2);margin-bottom:6px;">
        <span style="color:${color};font-weight:700;">${status} إجمالي المخصص: ${total} / ${qPerModel}</span>
      </div>
      ${bankFiles.map((f, i) => `
        <div style="background:var(--card);border:1px solid var(--border);border-radius:6px;
                    padding:10px;margin-bottom:8px;display:grid;
                    grid-template-columns:1fr 80px 50px;gap:8px;align-items:center;">
          <div>
            <div style="font-size:13px;font-weight:600;color:var(--text);">${f.name}</div>
            <div style="font-size:11px;color:var(--text3);">${f.questions.length} سؤال متاح</div>
          </div>
          <input type="number" min="0" max="${f.questions.length}" value="${f.alloc}"
                 onchange="updateBankFileAlloc(${i}, this.value)"
                 style="padding:6px;border:1px solid var(--border);border-radius:4px;
                        background:var(--bg3);color:var(--text);font-family:inherit;font-size:13px;">
          <button onclick="removeBankFile(${i})"
                  style="padding:6px 10px;background:var(--danger);color:#fff;border:none;
                         border-radius:4px;font-family:inherit;cursor:pointer;font-size:11px;">✕</button>
        </div>
      `).join('')}
    </div>`;
}

function mergeQuestionsFromBank() {
  if (!multiFileMode || bankFiles.length === 0) return questions;
  
  const merged = [];
  bankFiles.forEach(f => {
    const sample = shuffle(f.questions).slice(0, f.alloc);
    merged.push(...sample);
  });
  
  // If still short, fill from last file
  const qPerModel = parseInt(document.getElementById('qPerModel')?.value || 20);
  const n = parseInt(document.getElementById('numModels')?.value || 4);
  const needed = qPerModel * n;
  
  if (merged.length < needed && bankFiles.length > 0) {
    const lastFile = bankFiles[bankFiles.length - 1];
    const remaining = lastFile.questions.filter(q =>
      !merged.some(m => m[colMap.q] === q[colMap.q])
    );
    const toAdd = Math.min(remaining.length, needed - merged.length);
    merged.push(...remaining.slice(0, toAdd));
  }
  
  return merged;
}

// ══════════════════════════════════════════
// MULTI-PART MODE (Panel 1 — settings)
// Toggle-driven: split questions fairly across multiple files
// ══════════════════════════════════════════
let multiPartEnabled = false;
let mpParts = [];   // [{ name, questions, requested, color, colMap }]

const MP_COLORS = ['#4f7cff','#7c5cfc','#00d4aa','#ffb84f','#ff4f6a','#06b6d4','#a855f7','#f97316'];

function toggleMultiPartMode(el) {
  el.classList.toggle('on');
  multiPartEnabled = el.classList.contains('on');
  const card = document.getElementById('multipart-card');
  if (card) card.style.display = multiPartEnabled ? 'block' : 'none';
  if (multiPartEnabled) renderMpPartsList();
}

function handleMpFile(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb  = XLSX.read(e.target.result, { type: 'array' });
      const ws  = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
      if (raw.length < 2) { showToast('الملف فارغ', 'error'); return; }

      const headers  = raw[0].map(h => String(h||'').trim());
      const savedMap = { ...colMap };
      autoMap(headers);
      const fileColMap = { ...colMap };
      colMap = savedMap;

      const qs = raw.slice(1)
        .filter(r => r.some(c => c !== undefined && c !== ''))
        .map(r => {
          const obj = {};
          headers.forEach((h, i) => obj[h] = r[i] !== undefined ? String(r[i]).trim() : '');
          obj.__colMap = fileColMap;
          return obj;
        });

      const partName = file.name.replace(/\.[^.]+$/, '');
      const color    = MP_COLORS[mpParts.length % MP_COLORS.length];
      const qPerModel = parseInt(document.getElementById('qPerModel')?.value || 20);

      // Default requested: distribute evenly across parts (rough equal split suggestion)
      const suggested = Math.max(1, Math.floor(qPerModel / (mpParts.length + 1)));
      mpParts.push({ name: partName, questions: qs, requested: Math.min(suggested, qs.length), color, colMap: fileColMap });

      document.getElementById('mpFileInput').value = '';
      renderMpPartsList();
      showToast(`✅ تم إضافة "${partName}" — ${qs.length} سؤال`, 'success');
    } catch(err) {
      showToast('خطأ في قراءة الملف: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function removeMpPart(i) {
  mpParts.splice(i, 1);
  renderMpPartsList();
}

function updateMpPartCount(i, val) {
  mpParts[i].requested = Math.min(parseInt(val)||0, mpParts[i].questions.length);
  renderMpDistSummary();
}

function renderMpPartsList() {
  const container = document.getElementById('mp-parts-list');
  const summary   = document.getElementById('mp-dist-summary');
  if (!mpParts.length) { container.innerHTML = ''; if(summary) summary.style.display='none'; return; }

  container.innerHTML = mpParts.map((p, i) => `
    <div style="display:flex;align-items:center;gap:10px;padding:10px 12px;margin-bottom:8px;
                background:var(--bg3);border-radius:var(--radius-sm);
                border:1px solid var(--border);border-right:3px solid ${p.color};">
      <div style="flex:1;min-width:0;">
        <div style="font-size:13px;font-weight:600;color:var(--text);white-space:nowrap;
                    overflow:hidden;text-overflow:ellipsis;">${p.name}</div>
        <div style="font-size:11px;color:var(--text3);margin-top:2px;">
          متاح: <strong style="color:${p.color};">${p.questions.length}</strong> سؤال
        </div>
      </div>
      <div style="display:flex;align-items:center;gap:8px;flex-shrink:0;">
        <label style="font-size:11px;color:var(--text2);white-space:nowrap;">عدد الأسئلة:</label>
        <input type="number" min="0" max="${p.questions.length}"
               value="${p.requested || 0}"
               onchange="updateMpPartCount(${i}, this.value)"
               style="width:70px;text-align:center;padding:6px;border-radius:6px;
                      border:1px solid ${p.color};background:var(--bg);
                      color:var(--text);font-family:'Cairo',sans-serif;font-size:13px;">
      </div>
      <button onclick="removeMpPart(${i})"
              style="width:28px;height:28px;border-radius:6px;border:none;
                     background:#ff4f6a22;color:var(--danger);cursor:pointer;
                     font-size:14px;display:flex;align-items:center;justify-content:center;flex-shrink:0;">✕</button>
    </div>`).join('');

  if(summary) summary.style.display = 'block';
  renderMpDistSummary();

  // Setup drag-drop for mpDropZone
  const dz = document.getElementById('mpDropZone');
  if (dz && !dz._mpDragSetup) {
    dz._mpDragSetup = true;
    dz.addEventListener('dragover',  e => { e.preventDefault(); dz.classList.add('drag'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag'));
    dz.addEventListener('drop', e => {
      e.preventDefault(); dz.classList.remove('drag');
      handleMpFile({ target: { files: e.dataTransfer.files } });
    });
  }
}

function renderMpDistSummary() {
  const total     = mpParts.reduce((s, p) => s + (p.requested||0), 0);
  const qPerModel = parseInt(document.getElementById('qPerModel')?.value||20);
  const barsEl    = document.getElementById('mp-dist-bars');
  const warnEl    = document.getElementById('mp-dist-warning');
  if (!barsEl) return;

  barsEl.innerHTML = mpParts.map(p => {
    const pct = total > 0 ? Math.round((p.requested/total)*100) : 0;
    return `
      <div style="margin-bottom:6px;">
        <div style="display:flex;justify-content:space-between;font-size:11px;
                    color:var(--text2);margin-bottom:3px;">
          <span>${p.name}</span>
          <span style="color:${p.color};font-weight:600;">${p.requested} سؤال (${pct}%)</span>
        </div>
        <div style="height:6px;background:var(--border);border-radius:3px;overflow:hidden;">
          <div style="height:100%;width:${pct}%;background:${p.color};border-radius:3px;
                      transition:width .4s;"></div>
        </div>
      </div>`;
  }).join('') + `
    <div style="font-size:12px;color:var(--text2);margin-top:8px;padding-top:8px;
                border-top:1px solid var(--border);display:flex;justify-content:space-between;">
      <span>الإجمالي المخصص:</span>
      <strong style="color:${total===qPerModel?'var(--accent3)':total>qPerModel?'var(--danger)':'var(--warn)'};">
        ${total} / ${qPerModel}
      </strong>
    </div>`;

  if (warnEl) {
    if (total > qPerModel) {
      warnEl.style.display = 'block';
      warnEl.style.background = '#ff4f6a11';
      warnEl.style.border = '1px solid #ff4f6a44';
      warnEl.style.color = 'var(--danger)';
      warnEl.textContent = `⚠️ الإجمالي (${total}) يتجاوز عدد أسئلة النموذج (${qPerModel}). سيتم الاقتصار على ${qPerModel}.`;
    } else if (total < qPerModel && total > 0 && mpParts.length > 0) {
      warnEl.style.display = 'block';
      warnEl.style.background = '#ffb84f11';
      warnEl.style.border = '1px solid #ffb84f44';
      warnEl.style.color = 'var(--warn)';
      warnEl.textContent = `ℹ️ النقص (${qPerModel - total} سؤال) سيُكمَّل تلقائياً من آخر ملف مرفوع: "${mpParts[mpParts.length-1].name}".`;
    } else {
      warnEl.style.display = 'none';
    }
  }
}

function buildMpQuestionsPool() {
  if (!multiPartEnabled || mpParts.length === 0) return null;
  const qPerModel = parseInt(document.getElementById('qPerModel')?.value||20);
  const merged = [];

  for (let i = 0; i < mpParts.length; i++) {
    const p = mpParts[i];
    let take = p.requested || 0;
    take = Math.min(take, p.questions.length);
    const picked = shuffle([...p.questions]).slice(0, take).map(q => ({
      ...q,
      __partName:   p.name,
      __partColor:  p.color,
      __partColMap: p.colMap
    }));
    merged.push(...picked);
  }

  // Fill shortfall from last file
  if (merged.length < qPerModel && mpParts.length > 0) {
    const lastPart = mpParts[mpParts.length - 1];
    const usedTexts = new Set(merged.map(q => q[lastPart.colMap.q || Object.keys(q)[0]] || ''));
    const remaining = lastPart.questions.filter(q => {
      const key = lastPart.colMap.q || Object.keys(q)[0];
      return !usedTexts.has(q[key] || '');
    });
    const toAdd = Math.min(remaining.length, qPerModel - merged.length);
    const extra = shuffle(remaining).slice(0, toAdd).map(q => ({
      ...q,
      __partName:   lastPart.name,
      __partColor:  lastPart.color,
      __partColMap: lastPart.colMap
    }));
    merged.push(...extra);
  }

  // Use first part's colMap as default
  if (mpParts.length > 0) colMap = { ...mpParts[0].colMap };
  return merged;
}

// ══════════════════════════════════════════
// NAVIGATION
// ══════════════════════════════════════════
function goPanel(n) {
  if (n > 0 && n < 4 && questions.length === 0 && !multiPartEnabled) {
    showToast('يرجى استيراد الأسئلة أولاً', 'error');
    return;
  }
  if (n > 0 && n < 4 && multiPartEnabled && mpParts.length === 0) {
    showToast('يرجى إضافة أجزاء المادة في إعدادات النماذج', 'error');
    return;
  }
  document.querySelectorAll('.panel').forEach((p, i) => p.classList.toggle('active', i === n));
  document.querySelectorAll('.step-btn').forEach((b, i) => b.classList.toggle('active', i === n));

  // Initialize scanner panel when navigating to it
  if (n === 4) {
    initScannerPanel();
  }
  
  // Update UI elements if going to design panel
  if (n === 2) {
    const answerStyleSelect = document.getElementById('answerStyle');
    if (answerStyleSelect) {
      answerStyleSelect.value = design.answerStyle || 'normal';
    }
  }
  
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
  autoSaveCurrentProject();
}

function setAnswerTablePos(val) {
  opts.answerTablePos = val;
  const btnStart = document.getElementById('pos-start');
  const btnEnd   = document.getElementById('pos-end');
  if (!btnStart || !btnEnd) return;
  if (val === 'start') {
    btnStart.style.cssText = 'padding:5px 10px;border-radius:6px;font-family:Cairo,sans-serif;font-size:11px;font-weight:600;cursor:pointer;border:1px solid var(--accent);background:var(--accent);color:#fff;';
    btnEnd.style.cssText   = 'padding:5px 10px;border-radius:6px;font-family:Cairo,sans-serif;font-size:11px;font-weight:600;cursor:pointer;border:1px solid var(--border2);background:transparent;color:var(--text2);';
  } else {
    btnEnd.style.cssText   = 'padding:5px 10px;border-radius:6px;font-family:Cairo,sans-serif;font-size:11px;font-weight:600;cursor:pointer;border:1px solid var(--accent);background:var(--accent);color:#fff;';
    btnStart.style.cssText = 'padding:5px 10px;border-radius:6px;font-family:Cairo,sans-serif;font-size:11px;font-weight:600;cursor:pointer;border:1px solid var(--border2);background:transparent;color:var(--text2);';
  }
  autoSaveCurrentProject();
}

// ══════════════════════════════════════════
// FILE IMPORT
// ══════════════════════════════════════════
function initDropZone() {
  const dropZone = document.getElementById('dropZone');
  if (!dropZone) return;
  dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('drag'); });
  dropZone.addEventListener('dragleave', ()  => dropZone.classList.remove('drag'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag');
    handleFile({ target: { files: e.dataTransfer.files } });
  });

  const multiDrop = document.getElementById('multiDropZone');
  if (!multiDrop) return;
  multiDrop.addEventListener('dragover',  e => { e.preventDefault(); multiDrop.classList.add('drag'); });
  multiDrop.addEventListener('dragleave', ()  => multiDrop.classList.remove('drag'));
  multiDrop.addEventListener('drop', e => {
    e.preventDefault();
    multiDrop.classList.remove('drag');
    handleMultiFile({ target: { files: e.dataTransfer.files } });
  });
}

// ══════════════════════════════════════════
// IMPORT MODE — single / multi
// ══════════════════════════════════════════
let importMode = 'single';   // 'single' | 'multi'
let parts      = [];         // [{ name, questions, requested, color }]

const PART_COLORS = ['#4f7cff','#7c5cfc','#00d4aa','#ffb84f','#ff4f6a','#06b6d4','#a855f7','#f97316'];

function setImportMode(mode) {
  importMode = mode;
  const sBtn = document.getElementById('mode-single-btn');
  const mBtn = document.getElementById('mode-multi-btn');
  const sArea = document.getElementById('single-mode-area');
  const mArea = document.getElementById('multi-mode-area');

  if (mode === 'single') {
    sBtn.style.border = '2px solid var(--accent)';
    sBtn.style.background = 'var(--accent)11';
    sBtn.querySelector('div:nth-child(2)').style.color = 'var(--accent)';
    mBtn.style.border = '2px solid var(--border)';
    mBtn.style.background = 'transparent';
    mBtn.querySelector('div:nth-child(2)').style.color = 'var(--text2)';
    sArea.style.display = 'block';
    mArea.style.display = 'none';
  } else {
    mBtn.style.border = '2px solid var(--accent)';
    mBtn.style.background = 'var(--accent)11';
    mBtn.querySelector('div:nth-child(2)').style.color = 'var(--accent)';
    sBtn.style.border = '2px solid var(--border)';
    sBtn.style.background = 'transparent';
    sBtn.querySelector('div:nth-child(2)').style.color = 'var(--text2)';
    sArea.style.display = 'none';
    mArea.style.display = 'block';
  }
}

function resetImport() {
  questions = [];
  parts     = [];
  document.getElementById('preview-table').style.display = 'none';
  document.getElementById('stat-total').textContent  = '0';
  document.getElementById('stat-topics').textContent = '0';
  document.getElementById('stat-cols').textContent   = '0';
  document.getElementById('stat-ready').textContent  = '-';
  document.getElementById('sb-total').textContent    = '0';
  document.getElementById('sb-topics').textContent   = '0';
  renderPartsList();
}

// ── MULTI: handle new file dropped / selected ──
function handleMultiFile(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb  = XLSX.read(e.target.result, { type: 'array' });
      const ws  = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
      if (raw.length < 2) { showToast('الملف فارغ', 'error'); return; }

      const headers = raw[0].map(h => String(h||'').trim());
      // Temporarily map columns for this file
      const savedMap = { ...colMap };
      autoMap(headers);
      const fileColMap = { ...colMap };
      colMap = savedMap; // restore

      const qs = raw.slice(1)
        .filter(r => r.some(c => c !== undefined && c !== ''))
        .map(r => {
          const obj = {};
          headers.forEach((h, i) => obj[h] = r[i] !== undefined ? String(r[i]).trim() : '');
          obj.__colMap = fileColMap;  // embed mapping with each row
          return obj;
        });

      const partName = file.name.replace(/\.[^.]+$/, '');
      const color    = PART_COLORS[parts.length % PART_COLORS.length];
      parts.push({ name: partName, questions: qs, requested: 0, color, colMap: fileColMap });

      // Reset file input so same file can be re-added
      document.getElementById('multiFileInput').value = '';
      renderPartsList();
      showToast(`✅ تم إضافة "${partName}" — ${qs.length} سؤال`, 'success');
    } catch(err) {
      showToast('خطأ في قراءة الملف: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

// ── Render parts list with editable counts ──
function renderPartsList() {
  const container = document.getElementById('parts-list');
  const summary   = document.getElementById('dist-summary');
  if (!parts.length) { container.innerHTML = ''; summary.style.display = 'none'; return; }

  container.innerHTML = parts.map((p, i) => `
    <div style="display:flex;align-items:center;gap:10px;padding:10px 12px;margin-bottom:8px;
                background:var(--bg3);border-radius:var(--radius-sm);
                border:1px solid var(--border);border-right:3px solid ${p.color};">
      <div style="flex:1;min-width:0;">
        <div style="font-size:13px;font-weight:600;color:var(--text);white-space:nowrap;
                    overflow:hidden;text-overflow:ellipsis;">${p.name}</div>
        <div style="font-size:11px;color:var(--text3);margin-top:2px;">
          متاح: <strong style="color:${p.color};">${p.questions.length}</strong> سؤال
        </div>
      </div>
      <div style="display:flex;align-items:center;gap:8px;flex-shrink:0;">
        <label style="font-size:11px;color:var(--text2);white-space:nowrap;">عدد الأسئلة:</label>
        <input type="number" min="0" max="${p.questions.length}"
               value="${p.requested || 0}"
               onchange="updatePartCount(${i}, this.value)"
               style="width:70px;text-align:center;padding:6px;border-radius:6px;
                      border:1px solid ${p.color};background:var(--bg);
                      color:var(--text);font-family:'Cairo',sans-serif;font-size:13px;">
      </div>
      <button onclick="removePart(${i})"
              style="width:28px;height:28px;border-radius:6px;border:none;
                     background:#ff4f6a22;color:var(--danger);cursor:pointer;
                     font-size:14px;display:flex;align-items:center;justify-content:center;flex-shrink:0;">✕</button>
    </div>`).join('');

  summary.style.display = 'block';
  renderDistSummary();
}

function updatePartCount(i, val) {
  parts[i].requested = Math.min(parseInt(val)||0, parts[i].questions.length);
  renderDistSummary();
}

function removePart(i) {
  parts.splice(i, 1);
  renderPartsList();
}

function renderDistSummary() {
  const total     = parts.reduce((s, p) => s + (p.requested||0), 0);
  const qPerModel = parseInt(document.getElementById('qPerModel')?.value||20);
  const barsEl    = document.getElementById('dist-bars');
  const warnEl    = document.getElementById('dist-warning');

  barsEl.innerHTML = parts.map(p => {
    const pct = total > 0 ? Math.round((p.requested/total)*100) : 0;
    return `
      <div style="margin-bottom:6px;">
        <div style="display:flex;justify-content:space-between;font-size:11px;
                    color:var(--text2);margin-bottom:3px;">
          <span>${p.name}</span>
          <span style="color:${p.color};font-weight:600;">${p.requested} سؤال (${pct}%)</span>
        </div>
        <div style="height:6px;background:var(--border);border-radius:3px;overflow:hidden;">
          <div style="height:100%;width:${pct}%;background:${p.color};border-radius:3px;
                      transition:width .4s;"></div>
        </div>
      </div>`;
  }).join('') + `
    <div style="font-size:12px;color:var(--text2);margin-top:8px;padding-top:8px;
                border-top:1px solid var(--border);display:flex;justify-content:space-between;">
      <span>الإجمالي المطلوب:</span>
      <strong style="color:${total===qPerModel?'var(--accent3)':total>qPerModel?'var(--danger)':'var(--warn)'};">
        ${total} / ${qPerModel}
      </strong>
    </div>`;

  if (total > qPerModel) {
    warnEl.style.display = 'block';
    warnEl.textContent = `⚠️ الإجمالي (${total}) يتجاوز عدد أسئلة النموذج (${qPerModel}). سيتم الاقتصار على ${qPerModel}.`;
  } else if (total < qPerModel && total > 0) {
    warnEl.style.display = 'block';
    warnEl.style.background = '#ffb84f11';
    warnEl.style.borderColor = '#ffb84f44';
    warnEl.style.color = 'var(--warn)';
    warnEl.textContent = `ℹ️ النقص (${qPerModel - total} سؤال) سيُكمَّل تلقائياً من آخر جزء مرفوع.`;
  } else {
    warnEl.style.display = 'none';
  }

  // Merge multi-source questions into global questions array for preview
  if (total > 0) mergeAndPreviewMulti();
}

// ── Merge parts into global questions[] for preview and generation ──
function mergeAndPreviewMulti() {
  const merged = [];
  const qPerModel = parseInt(document.getElementById('qPerModel')?.value||20);
  let remaining = qPerModel;

  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    let take = p.requested || 0;
    // Last part fills the gap
    if (i === parts.length - 1 && merged.length < qPerModel) {
      take = Math.min(qPerModel - merged.length, p.questions.length);
    }
    take = Math.min(take, p.questions.length, remaining);
    // Embed colMap and part label into each question
    const picked = shuffle([...p.questions]).slice(0, take).map(q => ({
      ...q,
      __partName:   p.name,
      __partColor:  p.color,
      __partColMap: p.colMap
    }));
    merged.push(...picked);
    remaining -= take;
    if (remaining <= 0) break;
  }

  // Use the colMap of the first part as default (autoMap will handle per-question later)
  if (parts.length > 0) {
    colMap = { ...parts[0].colMap };
  }
  questions = merged;

  // Update stats
  document.getElementById('stat-total').textContent  = questions.length;
  document.getElementById('stat-topics').textContent = parts.length;
  document.getElementById('stat-cols').textContent   = '—';
  document.getElementById('stat-ready').textContent  = '✅';
  document.getElementById('sb-total').textContent    = questions.length;
  document.getElementById('sb-topics').textContent   = parts.length;
  document.getElementById('sb-avg').textContent      = '4';

  // Show preview
  renderMultiPreview();
}

function renderMultiPreview() {
  const preview = document.getElementById('preview-table');
  preview.style.display = 'block';
  document.getElementById('table-head').innerHTML =
    '<th>#</th><th>الجزء</th><th>السؤال</th><th>A</th><th>Answer</th>';
  document.getElementById('table-body').innerHTML =
    questions.slice(0, 10).map((q, i) => {
      const cm = q.__partColMap || colMap;
      return `<tr>
        <td>${i+1}</td>
        <td><span style="background:${q.__partColor||'var(--accent)'}22;
                         color:${q.__partColor||'var(--accent)'};
                         padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600;">
          ${q.__partName||'—'}
        </span></td>
        <td>${String(q[cm.q]||'').slice(0,50)}</td>
        <td>${String(q[cm.a]||'').slice(0,30)}</td>
        <td>${String(q[cm.ans]||'').slice(0,20)}</td>
      </tr>`;
    }).join('')
    + (questions.length > 10
      ? `<tr><td colspan="5" style="text-align:center;color:var(--text3)">... و ${questions.length-10} سؤال آخر</td></tr>`
      : '');
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
  autoSaveCurrentProject();
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

function updateSidebar() {
  const topics = new Set(questions.map(q => q[colMap.topic]).filter(Boolean));
  const sbTotal = document.getElementById('sb-total');
  const sbTopics = document.getElementById('sb-topics');
  const sbAvg = document.getElementById('sb-avg');
  if (sbTotal) sbTotal.textContent = questions.length;
  if (sbTopics) sbTopics.textContent = topics.size || 0;
  if (sbAvg) sbAvg.textContent = [colMap.a,colMap.b,colMap.c,colMap.d].filter(Boolean).length;
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
  autoSaveCurrentProject();
}

function setAnswerStyle() {
  const style = document.getElementById('answerStyle')?.value || 'normal';
  design.answerStyle = style;
  autoSaveCurrentProject();
  showToast(`تم تعيين نمط الإجابات: ${style === 'highlighted' ? 'تظليل تلقائي' : 'جدول عادي'} ✓`, 'success');
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
  
  // Update exam config from form
  examCfg.subject = g('subjectName')?.value || '';
  examCfg.institution = g('institution')?.value || '';
  examCfg.duration = g('examDuration')?.value || '';
  examCfg.grade = g('totalGrade')?.value || '';
  examCfg.date = g('examDate')?.value || '';
  examCfg.instructions = g('instructions')?.value || '';
  
  autoSaveCurrentProject();
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
  if (!questions.length && !multiPartEnabled) { showToast('لا توجد أسئلة للتوليد', 'error'); return; }
  if (multiPartEnabled && mpParts.length === 0) { showToast('يرجى إضافة أجزاء المادة أولاً', 'error'); return; }

  const n      = parseInt(document.getElementById('numModels').value)  || 4;
  const qCount = parseInt(document.getElementById('qPerModel').value)  || 20;

  if (!multiPartEnabled && qCount > questions.length) {
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

  // Use multi-part pool if enabled, else use main questions array
  const questionsForGen = multiPartEnabled ? buildMpQuestionsPool() : questions;
  if (!questionsForGen || questionsForGen.length === 0) { showToast('لا توجد أسئلة للتوليد', 'error'); return; }

  if (questionsForGen.length < qCount) {
    showToast(`عدد الأسئلة المتاح (${questionsForGen.length}) أقل من المطلوب (${qCount})`, 'error');
    return;
  }

  const btn = document.getElementById('gen-btn');
  btn.classList.add('generating');

  setTimeout(() => {
    models = [];
    const modelLetters = 'أبجدهوزحطيكلمنسعفصقرشتثخذضظغ'.split('');

    // Track globally used question indices for unique mode
    const globalUsedIndices = new Set();

    for (let m = 0; m < n; m++) {
      // Build candidate pool (by index to guarantee uniqueness)
      let pool = questionsForGen.map((q, idx) => ({ q, idx }));

      if (opts.unique) {
        const remaining = pool.filter(({ idx }) => !globalUsedIndices.has(idx));
        if (remaining.length >= qCount) pool = remaining;
        // else: wrap around — reuse all questions
      }

      // Shuffle and pick exactly qCount — NO duplicates within a model (each idx appears once)
      const selected = shuffle(pool).slice(0, qCount);
      selected.forEach(({ idx }) => globalUsedIndices.add(idx));

      // Process each question — use per-question colMap if available (multi-source mode)
      const processedQ = selected.map(({ q }) => {
        const cm = q.__partColMap || colMap;  // per-part column mapping
        const rawChoices = [
          { label: 'A', val: (q[cm.a] || '').trim() },
          { label: 'B', val: (q[cm.b] || '').trim() },
          { label: 'C', val: (q[cm.c] || '').trim() },
          { label: 'D', val: (q[cm.d] || '').trim() }
        ];

        // Remove blank + duplicate choice texts
        const uniqueChoices = deduplicateChoices(rawChoices);

        // Resolve correct answer BEFORE any shuffling (from original label or value)
        const correctRaw = (q[cm.ans] || '').trim();
        const correctOriginal = uniqueChoices.find(c =>
          c.val === correctRaw ||
          c.label === correctRaw ||
          c.label === mapEnToAr(correctRaw)
        );
        const correctVal = correctOriginal ? correctOriginal.val : correctRaw;

        // Shuffle VALUES only — labels stay A B C D in order
        let finalChoices;
        if (opts.shuffleA && uniqueChoices.length > 1) {
          const labels = uniqueChoices.map(c => c.label);
          const vals   = shuffle(uniqueChoices.map(c => c.val));
          finalChoices = labels.map((label, i) => ({ label, val: vals[i] }));
        } else {
          finalChoices = uniqueChoices;
        }

        // Find which label now holds the correct value
        const correctLabel = finalChoices.find(c => c.val === correctVal)?.label || correctRaw;

        return {
          text:         (q[cm.q] || '').trim(),
          choices:      finalChoices,
          correctLabel,
          topic:        (q[cm.topic] || q.__partName || '').trim(),
          diff:         (q[cm.diff]  || '').trim()
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
    autoSaveCurrentProject();
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
  const isOMR = design.answerStyle === 'highlighted';

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

  let omrGrid = '';
  let questionsHtml = '';
  let answerTable;
  
  // Always generate questions
  questionsHtml = model.questions.map((item, i) => {
    const maxLen = Math.max(...item.choices.map(c => (c.val || '').length));
    const colsPerRow = maxLen > 50 ? 2 : 4;
    const rows = [];
    for (let r = 0; r < item.choices.length; r += colsPerRow) {
      const cells = [];
      for (let c = r; c < Math.min(r + colsPerRow, item.choices.length); c++) {
        const choice = item.choices[c];
        const w = 100 / colsPerRow;
        cells.push(`<div style="width:${w}%;display:flex;gap:6px;align-items:flex-start;font-size:${fs-1}px;color:#4a5568;padding-right:8px;"><span style="font-weight:700;color:#2d3748;flex-shrink:0;min-width:18px;">${choice.label})</span><span>${renderMath(choice.val)}</span></div>`);
      }
      rows.push(`<div style="display:flex;width:100%;margin-bottom:6px;">${cells.join('')}</div>`);
    }
    const optHtml = rows.join('');
    return `
      <div style="display:block;width:100%;margin:0 0 1.4rem 0;padding-bottom:1rem;
                  border-bottom:1px solid #e0e4ef;page-break-inside:avoid;">
        <div style="font-size:${fs}px;line-height:1.9;margin-bottom:8px;color:#1a1a2e;font-weight:500;">
          <span style="font-weight:700;">${i+1})</span> ${renderMath(item.text)}
        </div>
        <div style="padding-right:24px;">${optHtml}</div>
      </div>`;
  }).join('');
  
  if (isOMR && opts.answerTable) {
    // OMR Grid format (empty) + questions - only if answerTable is enabled
    const omrCss = `.omr-circle{display:inline-block;width:11px;height:11px;border:1px solid #333;border-radius:50%;margin:0 auto;background:#fff;}.omr-table{width:100%;border-collapse:collapse;border:1px solid #666;margin-bottom:1rem;page-break-inside:avoid;font-size:7pt;}.omr-table td{border:1px solid #666;padding:2px 1px;text-align:center;font-weight:600;height:16px;}.omr-table .row-label{background:#e5e5e5;font-weight:700;color:#333;min-width:22px;}.omr-table .qnum-cell{background:#f0f0f0;font-weight:700;color:#333;font-size:6.5pt;}`;
    
    const tables = [];
    for (let tableStart = 0; tableStart < model.questions.length; tableStart += 20) {
      const tableQuestions = model.questions.slice(tableStart, Math.min(tableStart + 20, model.questions.length));
      const numCols = tableQuestions.length;
      
      let tableHtml = `<table class="omr-table">`;
      
      // Row 1: Question numbers
      tableHtml += `<tr><td class="row-label">Q</td>`;
      for (let i = 0; i < numCols; i++) {
        tableHtml += `<td class="qnum-cell">${tableStart + i + 1}</td>`;
      }
      tableHtml += `</tr>`;
      
      // Rows 2-5: Use actual answer labels from first question's choices
      const answerLabels = model.questions[0].choices.map(c => c.label) || ['A', 'B', 'C', 'D'];
      
      answerLabels.slice(0, 4).forEach((label, idx) => {
        tableHtml += `<tr><td class="row-label">${label}</td>`;
        for (let i = 0; i < numCols; i++) {
          tableHtml += `<td style="padding:1px;"><div class="omr-circle"></div></td>`;
        }
        tableHtml += `</tr>`;
      });
      
      tableHtml += `</table>`;
      tables.push(tableHtml);
    }
    
    omrGrid = `<style>${omrCss}</style><div style="padding:8px 12px;background:#f3f4f6;border:1px solid #d1d5db;border-radius:6px;margin-bottom:1rem;font-size:${fs-2}px;color:#374151;line-height:1.5;">
      <strong>ورقة الإجابات:</strong> ضع دائرة حول الإجابة الصحيحة
    </div>${tables.join('')}`;
    answerTable = '';
  } else {
    answerTable = opts.answerTable ? buildAnswerFillTable(model.questions.length, hBg, fs) : '';
  }
  
  const tableAtStart = opts.answerTablePos === 'start';
  const dir = detectDir();

  const paperHtml = `
    <div class="exam-paper" id="paper-${idx}" style="direction:${dir};">
      ${header}${student}${instructions}
      ${tableAtStart && (!isOMR || !opts.answerTable) ? answerTable : ''}
      ${tableAtStart && isOMR && opts.answerTable ? omrGrid : ''}
      <div style="display:block;width:100%;">${questionsHtml}</div>
      ${tableAtStart ? '' : (isOMR && opts.answerTable ? omrGrid : answerTable)}
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
  const perRow = 20;
  let tableRows = '';
  for (let r = 0; r < count; r += perRow) {
    const nums = [];
    const blanks = [];
    for (let c = r; c < Math.min(r + perRow, count); c++) {
      nums.push(`<td style="text-align:center;border:1px solid #c8d0e0;padding:2px 1px;
                             min-width:28px;font-size:6.5pt;font-weight:700;color:#333;">${c+1}</td>`);
      blanks.push(`<td style="border:1px solid #c8d0e0;padding:0;min-width:28px;height:16px;"></td>`);
    }
    // label column on left
    tableRows += `
      <tr>
        <td style="border:1px solid #c8d0e0;padding:2px 4px;font-size:6.5pt;font-weight:700;
                   color:#fff;background:${hBg};white-space:nowrap;width:48px;">Q</td>
        ${nums.join('')}
      </tr>
      <tr>
        <td style="border:1px solid #c8d0e0;padding:2px 4px;font-size:6.5pt;font-weight:700;
                   color:#fff;background:${hBg};white-space:nowrap;width:48px;">A</td>
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
  const content = models.map((m, modelIdx) => {
    const htmlContent = buildAnswerSheetHtml(m);
    return `
      <div style="margin-bottom:2rem;border:1px solid #ddd;border-radius:8px;overflow:hidden;">
        <div style="background:${design.headerBg};color:${design.headerText};padding:12px 16px;font-weight:700;font-size:14px;">
          📋 ورقة الإجابات — ${m.name}
        </div>
        <iframe id="sheet-${modelIdx}" style="width:100%;height:600px;border:none;background:#fff;" srcdoc="${htmlContent.replace(/"/g, '&quot;')}"></iframe>
      </div>`;
  }).join('');
  document.getElementById('answer-keys-content').innerHTML = content;
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
  const isOMR = design.answerStyle === 'highlighted';

  const dir = detectDir();
  
  // Always generate questions
  let questionsHtml = model.questions.map((item, i) => {
    const maxLen = Math.max(...item.choices.map(c => (c.val || '').length));
    const colsPerRow = maxLen > 50 ? 2 : 4;
    const width = 100 / colsPerRow;
    
    const pairs = [];
    for (let r = 0; r < item.choices.length; r += colsPerRow) {
      const cells = [];
      for (let c = r; c < Math.min(r + colsPerRow, item.choices.length); c++) {
        const choice = item.choices[c];
        cells.push(`<td style="width:${width}%;padding:2px 8px;font-size:9.5pt;color:#333;vertical-align:top;">
          ${choice ? `<b>${choice.label})</b> ${renderMath(choice.val)}` : ''}
        </td>`);
      }
      pairs.push(`<tr>${cells.join('')}</tr>`);
    }
    return `<div class="q-block" style="margin-bottom:9px;padding-bottom:8px;border-bottom:1px solid #dde2ee;">
      <div style="font-size:10.5pt;line-height:1.75;color:#1a1a2e;margin-bottom:4px;">
        <b>${i+1})</b> ${renderMath(item.text)}
      </div>
      <table style="width:100%;border-collapse:collapse;">${pairs.join('')}</table>
    </div>`;
  }).join('');
  
  let omrGrid = '';
  let fillTableHtml = '';
  
  if (isOMR && opts.answerTable) {
    // OMR Grid format (empty) - only if answerTable is enabled
    const omrCss = `.omr-circle{display:inline-block;width:11px;height:11px;border:1px solid #333;border-radius:50%;margin:0 auto;background:#fff;}.omr-table{width:100%;border-collapse:collapse;border:1px solid #666;margin-bottom:10pt;page-break-inside:avoid;font-size:7pt;}.omr-table td{border:1px solid #666;padding:2px 1px;text-align:center;font-weight:600;height:15px;}.omr-table .row-label{background:#e5e5e5;font-weight:700;color:#333;min-width:20px;}.omr-table .qnum-cell{background:#f0f0f0;font-weight:700;color:#333;font-size:6.5pt;}`;
    
    const tables = [];
    for (let tableStart = 0; tableStart < model.questions.length; tableStart += 20) {
      const tableQuestions = model.questions.slice(tableStart, Math.min(tableStart + 20, model.questions.length));
      const numCols = tableQuestions.length;
      
      let tableHtml = `<table class="omr-table">`;
      
      // Row 1: Question numbers
      tableHtml += `<tr><td class="row-label">Q</td>`;
      for (let i = 0; i < numCols; i++) {
        tableHtml += `<td class="qnum-cell">${tableStart + i + 1}</td>`;
      }
      tableHtml += `</tr>`;
      
      // Rows 2-5: Use actual answer labels from first question's choices
      const answerLabels = model.questions[0].choices.map(c => c.label) || ['A', 'B', 'C', 'D'];
      
      answerLabels.slice(0, 4).forEach((label, idx) => {
        tableHtml += `<tr><td class="row-label">${label}</td>`;
        for (let i = 0; i < numCols; i++) {
          tableHtml += `<td style="padding:1px;"><div class="omr-circle"></div></td>`;
        }
        tableHtml += `</tr>`;
      });
      
      tableHtml += `</table>`;
      tables.push(tableHtml);
    }
    
    omrGrid = `<style>${omrCss}</style><div style="padding:7px 10px;background:#f3f4f6;border:1px solid #d1d5db;border-radius:4px;margin-bottom:10pt;font-size:8pt;color:#374151;line-height:1.4;">
      <strong>ورقة الإجابات:</strong> ضع دائرة حول الإجابة الصحيحة
    </div>${tables.join('')}`;
  } else {
    // Fill table for normal format
    if (opts.answerTable) {
      const perRow = 20;
      let tRows = '';
      for (let r = 0; r < model.questions.length; r += perRow) {
        const nums = [], blanks = [];
        for (let c = r; c < Math.min(r+perRow, model.questions.length); c++) {
          nums.push(`<td style="text-align:center;border:1px solid #c8d0e0;padding:2px 1px;
                                 min-width:28px;font-size:6.5pt;font-weight:700;color:#333;">${c+1}</td>`);
          blanks.push(`<td style="border:1px solid #c8d0e0;min-width:28px;height:16px;"></td>`);
        }
        tRows += `<tr>
          <td style="border:1px solid #c8d0e0;padding:2px 4px;font-size:6.5pt;font-weight:700;
                     color:#fff;background:${hBg};white-space:nowrap;width:48px;">Q</td>${nums.join('')}
        </tr><tr>
          <td style="border:1px solid #c8d0e0;padding:2px 4px;font-size:6.5pt;font-weight:700;
                     color:#fff;background:${hBg};white-space:nowrap;width:48px;">A</td>${blanks.join('')}
        </tr>`;
      }
      fillTableHtml = `<div style="margin-top:14pt;page-break-inside:avoid;direction:ltr;">
        <table style="border-collapse:collapse;border:1px solid #c8d0e0;">${tRows}</table>
      </div>`;
    }
  }

  const tableAtStart = opts.answerTablePos === 'start';
  const qHtml = tableAtStart && isOMR && opts.answerTable ? omrGrid + questionsHtml : questionsHtml;
  const endHtml = !tableAtStart && isOMR && opts.answerTable ? omrGrid : fillTableHtml;

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
  ${tableAtStart && (!isOMR || !opts.answerTable) ? fillTableHtml : ''}
  ${qHtml}
  ${tableAtStart ? '' : endHtml}
</body></html>`;
}

// ══════════════════════════════════════════
// BUILD ANSWER SHEET HTML (A4)
// OMR Grid Style: Question numbers + circles for A B C D
// ══════════════════════════════════════════
function buildAnswerSheetHtml(model) {
  const hBg = design.headerBg, hTx = design.headerText;
  const isOMR = design.answerStyle === 'highlighted';

  if (!isOMR) {
    // Return plain table format (old version)
    const qHtml = model.questions.map((item, i) => {
      const maxLen = Math.max(...item.choices.map(c => (c.val || '').length));
      const colsPerRow = maxLen > 50 ? 2 : 4;
      const width = 100 / colsPerRow;
      
      const opts2 = item.choices.map(c => {
        return `<td style="width:${width}%;padding:2px 8px;font-size:9.5pt;vertical-align:top;color:#333;">
          <b style="flex-shrink:0;">${c.label})</b> ${renderMath(c.val)}</td>`;
      });
      
      const rows = [];
      for (let r = 0; r < opts2.length; r += colsPerRow) {
        const cells = [];
        for (let c = r; c < Math.min(r + colsPerRow, opts2.length); c++) {
          cells.push(opts2[c] || `<td style="width:${width}%;"></td>`);
        }
        rows.push(`<tr>${cells.join('')}</tr>`);
      }
      return `<div class="q-block" style="margin-bottom:9px;padding-bottom:8px;border-bottom:1px solid #dde2ee;">
        <div style="font-size:10.5pt;line-height:1.75;color:#1a1a2e;margin-bottom:4px;">
          <b>${i+1})</b> ${renderMath(item.text)}
        </div>
        <table style="width:100%;border-collapse:collapse;margin-right:16px;">${rows.join('')}</table>
      </div>`;
    }).join('');

    const keyCells = model.questions.map((q, i) => {
      return `<td style="text-align:center;border:1px solid #d1d5db;padding:5px 3px;min-width:40px;background:#f3f4f6;">
        <div style="font-size:8pt;color:#666;">${i+1}</div>
        <div style="font-size:12pt;font-weight:700;color:#374151;">${q.correctLabel}</div>
      </td>`;
    });
    const keyRows = [];
    for (let r = 0; r < keyCells.length; r += 20)
      keyRows.push(`<tr>${keyCells.slice(r, r+20).join('')}</tr>`);

    const dir = detectDir();
    return `<!DOCTYPE html><html lang="${dir==='ltr'?'en':'ar'}" dir="${dir}"><head><meta charset="UTF-8">
<style>${exportPageCss()}</style>
${mathJaxScript()}
</head><body style="direction:${dir};">
  <div style="background:${hBg};color:${hTx};padding:14px 20px;text-align:center;margin-bottom:12px;">
    <div style="font-size:13pt;font-weight:700;">${examCfg.institution||''}</div>
    <div style="font-size:17pt;font-weight:900;font-family:'Amiri',serif;margin:4px 0;">امتحان مادة: ${examCfg.subject||''}</div>
    <div style="font-size:10pt;opacity:.9;">الزمن: ${examCfg.duration} دقيقة | الدرجة: ${examCfg.grade} | ${examCfg.date||''}</div>
    <div style="margin-top:5px;display:inline-block;background:rgba(255,255,255,.25);padding:2px 18px;border-radius:20px;font-size:11pt;font-weight:700;">النموذج: ${model.letter}</div>
  </div>
  ${qHtml}
  <div style="margin-top:18pt;page-break-inside:avoid;">
    <div style="background:#4b5563;color:#fff;padding:5pt 12pt;font-size:10pt;font-weight:700;border-radius:4pt 4pt 0 0;display:inline-block;">ملخص الإجابات الصحيحة</div>
    <table style="width:100%;border-collapse:collapse;border:1px solid #d1d5db;">${keyRows.join('')}</table>
  </div>
</body></html>`;
  }

  // OMR Grid format
  const omrCss = `
    .omr-circle {
      display: inline-block;
      width: 14px;
      height: 14px;
      border: 1.5px solid #333;
      border-radius: 50%;
      margin: 0 auto;
      background: #fff;
    }
    .omr-circle.filled {
      background: #333;
    }
    .omr-table {
      width: 100%;
      border-collapse: collapse;
      border: 1px solid #333;
      margin-bottom: 14pt;
      page-break-inside: avoid;
    }
    .omr-table td {
      border: 1px solid #333;
      padding: 6px;
      text-align: center;
      font-size: 9pt;
      font-weight: 600;
    }
    .omr-table .row-label {
      background: #f3f4f6;
      font-weight: 700;
      color: #1f2937;
      min-width: 30px;
    }
    .omr-table .qnum-cell {
      background: #f9fafb;
      font-weight: 700;
      color: #374151;
    }
  `;

  // Group questions by 20 per table
  const tables = [];
  for (let tableStart = 0; tableStart < model.questions.length; tableStart += 20) {
    const tableQuestions = model.questions.slice(tableStart, Math.min(tableStart + 20, model.questions.length));
    const numCols = tableQuestions.length;
    
    let tableHtml = `<table class="omr-table">`;
    
    // Row 1: Question numbers
    tableHtml += `<tr><td class="row-label">السؤال</td>`;
    for (let i = 0; i < numCols; i++) {
      tableHtml += `<td class="qnum-cell">${tableStart + i + 1}</td>`;
    }
    tableHtml += `</tr>`;
    
    // Rows 2-5: use actual choice labels from the model (A/B/C/D or أ/ب/ج/د etc.)
    const firstQ = tableQuestions[0];
    const choiceLabels = firstQ ? firstQ.choices.map(c => c.label) : ['A','B','C','D'];

    choiceLabels.forEach((rowLabel, idx) => {
      tableHtml += `<tr><td class="row-label">${rowLabel}</td>`;
      for (let i = 0; i < numCols; i++) {
        const question = tableQuestions[i];
        const choices  = question ? question.choices || [] : [];
        const choice   = choices[idx];
        const isCorrect = choice && choice.label === question.correctLabel;
        const circleClass = isCorrect ? 'omr-circle filled' : 'omr-circle';
        tableHtml += `<td style="padding:8px 4px;"><div class="${circleClass}"></div></td>`;
      }
      tableHtml += `</tr>`;
    });
    
    tableHtml += `</table>`;
    tables.push(tableHtml);
  }

  // Summary grid
  const keyCells = model.questions.map((q, i) => `
    <td style="text-align:center;border:1px solid #d1d5db;padding:8px 4px;min-width:40px;background:#f9fafb;vertical-align:middle;">
      <div style="font-size:8pt;color:#666;margin-bottom:2px;">${i+1}</div>
      <div style="font-size:12pt;font-weight:700;color:#1f2937;">${q.correctLabel}</div>
    </td>`);
  const keyRows = [];
  for (let r = 0; r < keyCells.length; r += 20)
    keyRows.push(`<tr>${keyCells.slice(r, r+20).join('')}</tr>`);

  const dir = detectDir();
  
  return `<!DOCTYPE html><html lang="${dir==='ltr'?'en':'ar'}" dir="${dir}"><head><meta charset="UTF-8">
<style>${exportPageCss(omrCss)}</style>
${mathJaxScript()}
</head><body style="direction:${dir};">
  <div style="background:${hBg};color:${hTx};padding:14px 20px;text-align:center;margin-bottom:12px;">
    <div style="font-size:13pt;font-weight:700;">${examCfg.institution||''}</div>
    <div style="font-size:17pt;font-weight:900;font-family:'Amiri',serif;margin:4px 0;">امتحان مادة: ${examCfg.subject||''}</div>
    <div style="font-size:10pt;opacity:.9;">الزمن: ${examCfg.duration} دقيقة | الدرجة: ${examCfg.grade} | ${examCfg.date||''}</div>
    <div style="margin-top:5px;display:inline-flex;gap:8px;align-items:center;justify-content:center;flex-wrap:wrap;">
      <span style="display:inline-block;background:rgba(255,255,255,.25);padding:2px 18px;border-radius:20px;font-size:11pt;font-weight:700;">النموذج: ${model.letter}</span>
      <span style="display:inline-block;background:#1f2937;color:#fff;padding:2px 18px;border-radius:20px;font-size:10pt;font-weight:700;">جدول OMR</span>
    </div>
  </div>
  
  <div style="padding:10px 15px;background:#f3f4f6;border:1px solid #d1d5db;border-radius:6px;margin-bottom:12px;font-size:9pt;color:#374151;line-height:1.6;">
    <strong>التعليمات:</strong> املأ الدائرة المطابقة لكل إجابة صحيحة. الدوائر المملوءة تمثل الإجابات الصحيحة.
  </div>
  
  ${tables.join('')}
  
  <div style="margin-top:18pt;page-break-inside:avoid;">
    <div style="background:#1f2937;color:#fff;padding:8pt 12pt;font-size:10pt;font-weight:700;border-radius:4pt 4pt 0 0;display:inline-block;">✓ مفتاح الإجابات الصحيحة</div>
    <table style="width:100%;border-collapse:collapse;border:1px solid #d1d5db;">${keyRows.join('')}</table>
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
      // Check if any answer is too long
      const maxLen = Math.max(...item.choices.map(c => (c.val || '').length));
      const colsPerRow = maxLen > 50 ? 2 : 4;  // Use 2 cols only if any answer > 50 chars
      const width = 100 / colsPerRow;
      const pairs = [];
      for (let r = 0; r < item.choices.length; r += colsPerRow) {
        const cells = [];
        for (let c = r; c < Math.min(r + colsPerRow, item.choices.length); c++) {
          const choice = item.choices[c];
          cells.push(`<td width="${width}%" style="padding:2pt 6pt;font-size:10pt;">${choice ? `<b>${choice.label})</b> ${choice.val}` : ''}</td>`);
        }
        pairs.push(`<tr>${cells.join('')}</tr>`);
      }
      return `<div dir="${dir}" style="margin-bottom:6pt;">
        <p style="margin:0 0 4pt;font-size:10.5pt;line-height:1.75;"><b>${i+1})</b> ${item.text}</p>
        <table width="100%" dir="${dir}" style="border-collapse:collapse;margin-bottom:6pt;">${pairs.join('')}</table>
        <hr style="border:none;border-top:1px solid #dde;margin:0 0 4pt;">
      </div>`;
    }).join('');

    const isOMR = design.answerStyle === 'highlighted';
    let fillTable = '';
    if (opts.answerTable) {
      const perRow = 20;
      const tablePos = opts.answerTablePos === 'start' ? 'start' : 'end';

      if (isOMR) {
        // ── OMR circles table for Word export — full A4 width ──
        // A4 usable width with margins 20mm each side = ~170mm. Label col ~10mm, remaining /numCols.
        const choiceLabels = m.questions[0]?.choices?.map(c => c.label) || ['A','B','C','D'];
        const LABEL_W = 10;   // mm
        const PAGE_W  = 170;  // mm usable
        let omrTables = '';
        for (let tableStart = 0; tableStart < m.questions.length; tableStart += perRow) {
          const tableQuestions = m.questions.slice(tableStart, Math.min(tableStart + perRow, m.questions.length));
          const numCols = tableQuestions.length;
          const cellW   = ((PAGE_W - LABEL_W) / numCols).toFixed(1); // mm per question column

          let tHtml = `<table dir="ltr" style="border-collapse:collapse;border:1pt solid #555;margin-bottom:6pt;direction:ltr;width:100%;table-layout:fixed;">`;

          // Question numbers row
          tHtml += `<tr>`;
          tHtml += `<td style="border:1pt solid #555;padding:3pt 4pt;font-size:7.5pt;font-weight:700;background:#d0d4de;text-align:center;white-space:nowrap;width:${LABEL_W}mm;">Q</td>`;
          for (let i = 0; i < numCols; i++) {
            tHtml += `<td style="border:1pt solid #555;padding:3pt 1pt;font-size:7pt;font-weight:700;text-align:center;background:#e8eaf0;width:${cellW}mm;">${tableStart + i + 1}</td>`;
          }
          tHtml += `</tr>`;

          // Answer label rows with circles (○)
          choiceLabels.slice(0, 4).forEach(label => {
            tHtml += `<tr>`;
            tHtml += `<td style="border:1pt solid #555;padding:4pt 4pt;font-size:7.5pt;font-weight:700;background:#d0d4de;text-align:center;white-space:nowrap;">${label}</td>`;
            for (let i = 0; i < numCols; i++) {
              tHtml += `<td style="border:1pt solid #555;padding:4pt 1pt;text-align:center;font-size:11pt;line-height:1;height:16pt;">&#9711;</td>`;
            }
            tHtml += `</tr>`;
          });
          tHtml += `</table>`;
          omrTables += tHtml;
        }
        const omrLabel = `<p style="font-size:8.5pt;color:#374151;margin-bottom:5pt;"><strong>ورقة الإجابات:</strong> ضع دائرة حول الإجابة الصحيحة</p>`;
        fillTable = `<br>${omrLabel}${omrTables}`;
      } else {
        // ── Standard fill table ──
        let tRows = '';
        for (let r = 0; r < m.questions.length; r += perRow) {
          const nums = [], blanks = [];
          for (let c = r; c < Math.min(r+perRow, m.questions.length); c++) {
            nums.push(`<td style="text-align:center;border:1pt solid #c8d0e0;padding:2pt;width:20pt;font-size:7pt;font-weight:700;">${c+1}</td>`);
            blanks.push(`<td style="border:1pt solid #c8d0e0;width:20pt;height:12pt;"></td>`);
          }
          tRows += `<tr>
            <td style="border:1pt solid #c8d0e0;padding:2pt 4pt;font-size:7pt;font-weight:700;
                       background:${hBg};color:#fff;white-space:nowrap;width:30pt;">Q</td>${nums.join('')}
          </tr><tr>
            <td style="border:1pt solid #c8d0e0;padding:2pt 4pt;font-size:7pt;font-weight:700;
                       background:${hBg};color:#fff;white-space:nowrap;width:30pt;">A</td>${blanks.join('')}
          </tr>`;
        }
        fillTable = `<br><table dir="ltr" style="border-collapse:collapse;direction:ltr;">${tRows}</table>`;
      }

      if (tablePos === 'start') fillTable = fillTable + '<!--FILLTABLE_START-->';
    }

    const wordBody = fillTable.includes('<!--FILLTABLE_START-->')
      ? fillTable.replace('<!--FILLTABLE_START-->','') + studentTable + instrBlock + qHtml
      : studentTable + instrBlock + qHtml + fillTable;
    qFolder.file(`${m.name} - أسئلة.doc`,
      '\ufeff' + wordWrap(wordBody, m.letter, false));

    // ── Answer sheet body ──
    const aHtml = m.questions.map((item, i) => {
      // Check if any answer is too long
      const maxLen = Math.max(...item.choices.map(c => (c.val || '').length));
      const colsPerRow = maxLen > 50 ? 2 : 4;  // Use 2 cols only if any answer > 50 chars
      const width = 100 / colsPerRow;
      const pairs = [];
      for (let r = 0; r < item.choices.length; r += colsPerRow) {
        const cells = [];
        for (let c = r; c < Math.min(r + colsPerRow, item.choices.length); c++) {
          const choice = item.choices[c];
          if (!choice) continue;
          const ok = choice.label === item.correctLabel;
          cells.push(`<td width="${width}%" style="padding:2pt 6pt;font-size:10pt;">
            <span${ok ? ' class="correct" style="background:#bbf7d0;padding:1pt 4pt;"' : ''}>
              <b${ok ? ' style="color:#15803d;"' : ''}>${choice.label})</b> ${choice.val}
            </span></td>`);
        }
        pairs.push(`<tr>${cells.join('')}</tr>`);
      }
      return `<div dir="${dir}" style="margin-bottom:6pt;">
        <p style="margin:0 0 4pt;font-size:10.5pt;line-height:1.75;"><b>${i+1})</b> ${item.text}</p>
        <table width="100%" dir="${dir}" style="border-collapse:collapse;margin-bottom:6pt;">${pairs.join('')}</table>
        <hr style="border:none;border-top:1px solid #dde;margin:0 0 4pt;">
      </div>`;
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
// PROJECT MANAGER UI
// ══════════════════════════════════════════
function showProjectManager() {
  const projects = getAllProjects();
  
  let projectsHtml = '';
  if (projects.length === 0) {
    projectsHtml = `
      <div style="text-align:center;padding:2rem;color:var(--text3);">
        <div style="font-size:48px;margin-bottom:1rem;">📭</div>
        <p>لا توجد مشاريع محفوظة</p>
      </div>`;
  } else {
    projectsHtml = `
      <div style="display:grid;gap:10px;">
        ${projects.map(p => `
          <div style="display:flex;align-items:center;justify-content:space-between;
                     padding:12px;background:var(--bg3);border-radius:8px;border:1px solid var(--border);">
            <div style="flex:1;min-width:0;">
              <div style="font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
                ${p.name}
              </div>
              <div style="font-size:12px;color:var(--text3);margin-top:4px;">
                ${p.questions} سؤال • ${new Date(p.date).toLocaleString('ar-EG')}
              </div>
            </div>
            <div style="display:flex;gap:6px;margin-right:1rem;">
              <button onclick="loadProject(${p.id})" style="padding:6px 12px;background:var(--accent);color:#fff;
                     border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600;">فتح</button>
              <button onclick="deleteProjectConfirm(${p.id})" style="padding:6px 12px;background:#ff4f6a44;
                     color:#ff4f6a;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600;">حذف</button>
            </div>
          </div>
        `).join('')}
      </div>`;
  }
  
  const html = `
    <div style="position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.5);
                display:flex;align-items:center;justify-content:center;z-index:9999;" id="manager-overlay" onclick="closeProjectManager()">
      <div style="background:var(--bg2);border-radius:12px;padding:2rem;max-width:600px;width:90%;max-height:80vh;
                  overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.3);" onclick="event.stopPropagation()">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:1.5rem;">
          <h2 style="font-size:20px;font-weight:700;color:var(--text);margin:0;">📁 إدارة المشاريع</h2>
          <button onclick="closeProjectManager()" style="background:transparent;border:none;font-size:24px;cursor:pointer;color:var(--text2);">✕</button>
        </div>
        <div style="background:var(--bg3);border:1px solid var(--border);border-radius:8px;padding:10px 12px;
                    margin-bottom:1rem;font-size:12px;color:var(--text2);">
          💾 <strong>موقع الحفظ:</strong> localStorage في المتصفح (آمن ومحلي)<br>
          📍 <strong>المسار:</strong> AppData/Local/[Browser]/User Data/Local Storage
        </div>
        ${projectsHtml}
      </div>
    </div>`;
  
  document.body.insertAdjacentHTML('beforeend', html);
}

function closeProjectManager() {
  const overlay = document.getElementById('manager-overlay');
  if (overlay) overlay.remove();
}

function loadProject(id) {
  const project = loadProjectFromStorage(id);
  if (project) {
    questions = project.questions || [];
    colMap = project.colMap || colMap;
    opts = project.opts || opts;
    design = project.design || design;
    examCfg = project.examCfg || {};
    models = project.models || [];
    window.currentProjectId = id;

    // ── Restore form fields so examCfg name is preserved on next auto-save ──
    const setVal = (id, val) => { const el = document.getElementById(id); if (el) el.value = val || ''; };
    setVal('subjectName',  examCfg.subject);
    setVal('institution',  examCfg.institution);
    setVal('examDuration', examCfg.duration);
    setVal('totalGrade',   examCfg.grade);
    setVal('classLevel',   examCfg.classLevel);
    setVal('examDate',     examCfg.date);
    setVal('instructions', examCfg.instructions);

    updateSidebar();
    if (questions.length > 0) {
      renderPreviewTable(Object.keys(colMap).filter(k => colMap[k]));
    }

    if (models.length > 0) {
      goPanel(3);
      setTimeout(() => {
        renderModelTabs();
        showToast(`تم تحميل "${examCfg.subject || 'المشروع'}" — ${models.length} نماذج ✓`, 'success');
      }, 100);
    } else {
      goPanel(0);
      showToast(`تم تحميل "${examCfg.subject || 'المشروع'}" بنجاح ✓`, 'success');
    }

    closeProjectManager();
  } else {
    showToast('خطأ في تحميل المشروع', 'error');
  }
}

function deleteProjectConfirm(id) {
  if (confirm('هل أنت متأكد من حذف هذا المشروع؟')) {
    deleteProjectFromStorage(id);
    showToast('تم حذف المشروع ✓', 'success');
    closeProjectManager();
    setTimeout(showProjectManager, 300);
  }
}

// ══════════════════════════════════════════
// INIT
// ══════════════════════════════════════════
// ══════════════════════════════════════════
// THEME TOGGLE
// ══════════════════════════════════════════
let currentTheme = 'dark';

function toggleTheme() {
  currentTheme = currentTheme === 'dark' ? 'light' : 'dark';
  applyTheme(currentTheme);
  const btn = document.getElementById('theme-btn');
  if (btn) btn.textContent = currentTheme === 'dark' ? '☀️ Light' : '🌙 Dark';
}

function applyTheme(theme) {
  const root = document.documentElement;
  if (theme === 'light') {
    root.style.setProperty('--bg',      '#f0f2f8');
    root.style.setProperty('--bg2',     '#e4e8f2');
    root.style.setProperty('--bg3',     '#d8ddf0');
    root.style.setProperty('--card',    '#ffffff');
    root.style.setProperty('--border',  '#c8d0e8');
    root.style.setProperty('--border2', '#a0aace');
    root.style.setProperty('--text',    '#1a1f35');
    root.style.setProperty('--text2',   '#4a5580');
    root.style.setProperty('--text3',   '#7a85aa');
  } else {
    root.style.setProperty('--bg',      '#0f1117');
    root.style.setProperty('--bg2',     '#161b27');
    root.style.setProperty('--bg3',     '#1e2535');
    root.style.setProperty('--card',    '#1a2030');
    root.style.setProperty('--border',  '#2a3347');
    root.style.setProperty('--border2', '#3a4a67');
    root.style.setProperty('--text',    '#e8edf8');
    root.style.setProperty('--text2',   '#8a9bc0');
    root.style.setProperty('--text3',   '#5a6a8a');
  }
}

// ══════════════════════════════════════════
// OMR SCANNER MODULE
// ══════════════════════════════════════════

const RESULTS_STORAGE_KEY = 'omr_results_v1';

let scannerState = {
  selectedSubject: null,
  selectedModelIdx: null,
  selectedModel: null,
  currentStudentId: null,
  currentScannedAnswers: null,   // {1: 'A', 2: null, ...}
  pendingManualQuestions: [],    // question indices needing manual input
  cameraStream: null,
  facingMode: 'environment',
  allResults: []                 // [{studentId, modelName, answers, score, total, subject}]
};

// ── Load saved results from localStorage ──
function loadScanResults() {
  try {
    const data = localStorage.getItem(RESULTS_STORAGE_KEY);
    return data ? JSON.parse(data) : [];
  } catch(e) { return []; }
}

// ── Save results to localStorage ──
function saveScanResults(results) {
  try {
    localStorage.setItem(RESULTS_STORAGE_KEY, JSON.stringify(results));
  } catch(e) { console.error('Save results error:', e); }
}

// ── Initialize scanner panel ──
function initScannerPanel() {
  // Populate subject dropdown from saved projects
  const projects = getAllProjects();
  const subjectSelect = document.getElementById('scanner-subject-select');
  if (!subjectSelect) return;

  subjectSelect.innerHTML = '<option value="">-- اختر المادة --</option>';

  // Add subjects from current session if models exist
  if (models.length > 0 && examCfg.subject) {
    const opt = document.createElement('option');
    opt.value = '__current__';
    opt.textContent = `${examCfg.subject} (الجلسة الحالية — ${models.length} نماذج)`;
    subjectSelect.appendChild(opt);
  }

  // Add subjects from saved projects
  projects.forEach(p => {
    const opt = document.createElement('option');
    opt.value = p.id;
    opt.textContent = `${p.name} (${p.models} نماذج — ${new Date(p.date).toLocaleDateString('ar-EG')})`;
    subjectSelect.appendChild(opt);
  });

  // Load and show results table if any results exist
  scannerState.allResults = loadScanResults();
  if (scannerState.allResults.length > 0) {
    renderResultsTable();
  }
}

// ── Subject selection changed ──
function onScannerSubjectChange() {
  const val = document.getElementById('scanner-subject-select').value;
  const modelSelect = document.getElementById('scanner-model-select');
  const studentCard = document.getElementById('scanner-student-card');
  const modelInfo = document.getElementById('scanner-model-info');

  modelSelect.innerHTML = '<option value="">-- اختر النموذج --</option>';
  studentCard.style.display = 'none';
  modelInfo.style.display = 'none';
  scannerState.selectedSubject = null;

  if (!val) return;

  let loadedModels = [];
  let subjectName = '';

  if (val === '__current__') {
    loadedModels = models;
    subjectName = examCfg.subject;
  } else {
    const project = loadProjectFromStorage(parseInt(val));
    if (project) {
      loadedModels = project.models || [];
      subjectName = project.examCfg?.subject || 'غير محدد';
    }
  }

  if (loadedModels.length === 0) {
    showToast('لا توجد نماذج لهذه المادة', 'error');
    return;
  }

  scannerState.selectedSubject = { val, models: loadedModels, name: subjectName };

  loadedModels.forEach((m, i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = `${m.name} — ${m.questions.length} سؤال`;
    modelSelect.appendChild(opt);
  });

  // Add "auto-detect" option
  const autoOpt = document.createElement('option');
  autoOpt.value = 'auto';
  autoOpt.textContent = '🔍 تحديد تلقائي (من الورقة)';
  modelSelect.insertBefore(autoOpt, modelSelect.children[1]);

  studentCard.style.display = 'block';
  document.getElementById('results-subject-label').textContent = subjectName;
}

// ── Model selection changed ──
function onScannerModelChange() {
  const val = document.getElementById('scanner-model-select').value;
  const modelInfo = document.getElementById('scanner-model-info');

  if (!val || !scannerState.selectedSubject) { modelInfo.style.display = 'none'; return; }

  if (val === 'auto') {
    modelInfo.style.display = 'block';
    modelInfo.innerHTML = '🔍 سيتم تحديد النموذج تلقائياً من الورقة. إذا لم يتمكن الماسح من تحديده، سيطلب منك الاختيار يدوياً.';
    scannerState.selectedModelIdx = 'auto';
    return;
  }

  const idx = parseInt(val);
  const m = scannerState.selectedSubject.models[idx];
  scannerState.selectedModelIdx = idx;
  scannerState.selectedModel = m;

  modelInfo.style.display = 'block';
  modelInfo.innerHTML = `✅ النموذج: <strong style="color:var(--accent);">${m.name}</strong> | عدد الأسئلة: <strong>${m.questions.length}</strong> | مفتاح الإجابات جاهز للمقارنة`;
}

// ── Start scan session for a student ──
function startScanSession() {
  const studentId = document.getElementById('scanner-student-id').value.trim();
  if (!studentId) { showToast('يرجى إدخال الرقم الجامعي', 'error'); return; }

  const modelVal = document.getElementById('scanner-model-select').value;
  if (!modelVal) { showToast('يرجى اختيار النموذج', 'error'); return; }

  // Check for duplicate
  const existing = scannerState.allResults.find(r =>
    r.studentId === studentId && r.subject === scannerState.selectedSubject?.name
  );
  if (existing) {
    if (!confirm(`الرقم الجامعي "${studentId}" موجود مسبقاً لهذه المادة. هل تريد الاستبدال؟`)) return;
    scannerState.allResults = scannerState.allResults.filter(r =>
      !(r.studentId === studentId && r.subject === scannerState.selectedSubject?.name)
    );
  }

  scannerState.currentStudentId = studentId;
  document.getElementById('scanner-current-student').textContent = studentId;
  document.getElementById('scanner-camera-card').style.display = 'block';
  document.getElementById('scan-result-content').textContent = 'اضغط على زر المسح لبدء العملية...';
  document.getElementById('manual-override-panel').style.display = 'none';
  document.getElementById('score-summary-panel').style.display = 'none';

  startCamera();
}

// ── Camera management ──
async function startCamera() {
  try {
    if (scannerState.cameraStream) {
      scannerState.cameraStream.getTracks().forEach(t => t.stop());
    }
    const stream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: scannerState.facingMode, width: { ideal: 1280 }, height: { ideal: 960 } }
    });
    scannerState.cameraStream = stream;
    const video = document.getElementById('scanner-video');
    video.srcObject = stream;
  } catch(e) {
    showToast('لا يمكن الوصول للكاميرا: ' + e.message, 'error');
  }
}

function stopCamera() {
  if (scannerState.cameraStream) {
    scannerState.cameraStream.getTracks().forEach(t => t.stop());
    scannerState.cameraStream = null;
  }
  const video = document.getElementById('scanner-video');
  video.srcObject = null;
}

function switchCamera() {
  scannerState.facingMode = scannerState.facingMode === 'environment' ? 'user' : 'environment';
  startCamera();
}

// ── Capture frame from camera and scan ──
function captureAndScan() {
  const video = document.getElementById('scanner-video');
  if (!video.srcObject) { showToast('الكاميرا غير مفعّلة', 'error'); return; }

  const canvas = document.createElement('canvas');
  canvas.width  = video.videoWidth  || 640;
  canvas.height = video.videoHeight || 480;
  const ctx = canvas.getContext('2d');
  ctx.drawImage(video, 0, 0, canvas.width, canvas.height);

  const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  processOMRImage(imageData, canvas.width, canvas.height);
}

// ── Scan from uploaded file ──
function scanFromFile(event) {
  const file = event.target.files[0];
  if (!file) return;
  const img = new Image();
  img.onload = () => {
    const canvas = document.createElement('canvas');
    canvas.width  = img.width;
    canvas.height = img.height;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0);
    const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    processOMRImage(imageData, canvas.width, canvas.height);
  };
  img.src = URL.createObjectURL(file);
}

// ── OMR Image Processing Engine ──
function processOMRImage(imageData, width, height) {
  const model = scannerState.selectedModel || (
    scannerState.selectedSubject?.models[0]
  );

  if (!model) {
    showToast('يرجى اختيار النموذج أولاً', 'error');
    return;
  }

  document.getElementById('scan-result-content').innerHTML =
    '<div style="color:var(--accent);">⏳ جارٍ تحليل الصورة...</div>';

  setTimeout(() => {
    try {
      const answers = analyzeOMRBubbles(imageData, width, height, model.questions.length);
      scannerState.currentScannedAnswers = answers;

      // Find unclear/multiple-answer questions
      const unclear = [];
      Object.entries(answers).forEach(([qNum, ans]) => {
        if (ans === null || ans === 'MULTI') {
          unclear.push(parseInt(qNum));
        }
      });
      scannerState.pendingManualQuestions = unclear;

      // Show scan results
      showScanResults(answers, model);

      if (unclear.length > 0) {
        showManualOverride(unclear, model);
      } else {
        computeAndShowScore(answers, model);
      }
    } catch(e) {
      document.getElementById('scan-result-content').innerHTML =
        `<div style="color:var(--danger);">❌ خطأ في تحليل الصورة: ${e.message}</div>`;
      console.error('OMR error:', e);
    }
  }, 100);
}

// ── Bubble analysis algorithm ──
// Layout matches the OMR table format shown in the design:
//   Rows (top→bottom): header row (Q numbers) + 4 answer rows (أ,ب,ج,د) = 5 rows
//   Columns (right→left in RTL, but pixels are LTR): label col + one col per question
//
// So for each question Q (1..N):
//   Column index = Q  (skipping col-0 which is the label)
//   We check 4 rows (1..4) and detect which is filled
//
// Grid auto-detected from dark-line projections.
function analyzeOMRBubbles(imageData, width, height, numQuestions) {
  const data = imageData.data;
  const answers = {};

  // ── Grayscale ──
  const gray = new Uint8Array(width * height);
  for (let i = 0; i < width * height; i++) {
    const r = data[i*4], g = data[i*4+1], b = data[i*4+2];
    gray[i] = Math.round(0.299*r + 0.587*g + 0.114*b);
  }

  // ── Dark-line projection to find grid bounds ──
  const DARK = 80;
  const rowDark = new Float32Array(height);
  for (let y = 0; y < height; y++) {
    let cnt = 0;
    for (let x = 0; x < width; x++) { if (gray[y*width+x] < DARK) cnt++; }
    rowDark[y] = cnt / width;
  }
  const colDark = new Float32Array(width);
  for (let x = 0; x < width; x++) {
    let cnt = 0;
    for (let y = 0; y < height; y++) { if (gray[y*width+x] < DARK) cnt++; }
    colDark[x] = cnt / height;
  }

  // Find grid top/bottom
  const ROW_T = 0.12;
  let gridTop = Math.floor(height*0.05), gridBottom = Math.floor(height*0.95);
  for (let y = Math.floor(height*0.05); y < Math.floor(height*0.95); y++) {
    if (rowDark[y] >= ROW_T) { gridTop = y; break; }
  }
  for (let y = Math.floor(height*0.95); y > gridTop; y--) {
    if (rowDark[y] >= ROW_T) { gridBottom = y; break; }
  }

  // Find grid left/right
  const COL_T = 0.08;
  let gridLeft = Math.floor(width*0.02), gridRight = Math.floor(width*0.98);
  for (let x = Math.floor(width*0.02); x < Math.floor(width*0.98); x++) {
    if (colDark[x] >= COL_T) { gridLeft = x; break; }
  }
  for (let x = Math.floor(width*0.98); x > gridLeft; x--) {
    if (colDark[x] >= COL_T) { gridRight = x; break; }
  }

  // ── Grid division ──
  // 5 rows (header + 4 answers), (numQuestions+1) cols (label + questions)
  const numRows = 5;
  const numCols = numQuestions + 1;
  const gridH   = gridBottom - gridTop;
  const gridW   = gridRight  - gridLeft;
  const rowH    = gridH / numRows;
  const colW    = gridW / numCols;

  const labels  = ['A','B','C','D']; // answer rows 1..4
  const FILL_DARK      = 110; // pixel below this = dark/filled
  const FILL_THRESHOLD = 0.15; // ≥15% dark = bubble filled

  for (let q = 0; q < numQuestions; q++) {
    // Column: skip col-0 (label column)
    const colIdx = q + 1;
    const cLeft  = Math.floor(gridLeft + colIdx * colW);
    const cRight = Math.floor(gridLeft + (colIdx+1) * colW);

    const fillRatios = [];

    for (let r = 0; r < 4; r++) {
      // Row: skip row-0 (header with question numbers)
      const rowIdx = r + 1;
      const rTop   = Math.floor(gridTop + rowIdx * rowH);
      const rBot   = Math.floor(gridTop + (rowIdx+1) * rowH);

      // Inner margin — avoid cell borders
      const mH = Math.max(2, Math.floor(rowH * 0.20));
      const mW = Math.max(2, Math.floor(colW * 0.20));

      let dark = 0, total = 0;
      for (let y = rTop+mH; y < rBot-mH; y++) {
        for (let x = cLeft+mW; x < cRight-mW; x++) {
          if (x >= 0 && x < width && y >= 0 && y < height) {
            if (gray[y*width+x] < FILL_DARK) dark++;
            total++;
          }
        }
      }
      fillRatios.push(total > 0 ? dark/total : 0);
    }

    // Determine answer
    const filled = fillRatios
      .map((ratio, i) => ({ label: labels[i], ratio }))
      .filter(x => x.ratio >= FILL_THRESHOLD);

    if (filled.length === 0) {
      answers[q+1] = null;
    } else if (filled.length === 1) {
      answers[q+1] = filled[0].label;
    } else {
      filled.sort((a,b) => b.ratio - a.ratio);
      const gap = filled[0].ratio - filled[1].ratio;
      answers[q+1] = gap < 0.07 ? 'MULTI' : filled[0].label;
    }
  }

  return answers;
}

// ── Show scan results in UI ──
function showScanResults(answers, model) {
  const totalQ = model.questions.length;
  const detected = Object.values(answers).filter(v => v && v !== 'MULTI').length;
  const unclear  = Object.values(answers).filter(v => v === null || v === 'MULTI').length;

  let html = `<div style="margin-bottom:10px;font-size:12px;color:var(--text2);">
    تم اكتشاف <strong style="color:var(--accent3);">${detected}</strong> إجابة من أصل ${totalQ}
    ${unclear > 0 ? `| <strong style="color:var(--warn);">${unclear}</strong> غير واضحة` : ''}
  </div>`;

  html += '<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:4px;">';
  for (let q = 1; q <= totalQ; q++) {
    const ans = answers[q];
    let bg = 'var(--bg)';
    let color = 'var(--text2)';
    let label = ans || '?';

    if (ans === null)    { bg = '#ffb84f22'; color = 'var(--warn)'; label = '؟'; }
    if (ans === 'MULTI') { bg = '#ff4f6a22'; color = 'var(--danger)'; label = '!!'; }
    if (ans && ans !== 'MULTI') { bg = '#4f7cff22'; color = 'var(--accent)'; }

    html += `<div style="padding:4px;background:${bg};border-radius:6px;text-align:center;border:1px solid ${bg};">
      <div style="font-size:9px;color:var(--text3);">${q}</div>
      <div style="font-size:13px;font-weight:700;color:${color};">${label}</div>
    </div>`;
  }
  html += '</div>';

  document.getElementById('scan-result-content').innerHTML = html;
}

// ── Show manual override panel for unclear questions ──
function showManualOverride(unclearList, model) {
  const panel = document.getElementById('manual-override-panel');
  const list  = document.getElementById('manual-override-list');

  list.innerHTML = unclearList.map(qNum => {
    const q = model.questions[qNum - 1];
    const choiceLabels = q ? q.choices.map(c => c.label) : ['A','B','C','D'];
    const currentAns   = scannerState.currentScannedAnswers[qNum];
    const isMulti      = currentAns === 'MULTI';

    return `<div style="padding:8px;background:var(--bg);border-radius:8px;margin-bottom:8px;border:1px solid var(--border);">
      <div style="font-size:12px;color:var(--text2);margin-bottom:6px;">
        <strong style="color:var(--warn);">س${qNum}:</strong>
        ${q ? q.text.slice(0, 60) + (q.text.length > 60 ? '...' : '') : 'سؤال رقم ' + qNum}
        ${isMulti ? '<span style="color:var(--danger);font-size:11px;"> (إجابات متعددة)</span>' : '<span style="color:var(--warn);font-size:11px;"> (لم تُكتشف)</span>'}
      </div>
      <div style="display:flex;gap:6px;flex-wrap:wrap;">
        ${choiceLabels.map(label =>
          `<button onclick="setManualAnswer(${qNum},'${label}',this)"
                  style="padding:6px 14px;border-radius:6px;border:1px solid var(--border2);
                         background:var(--bg3);color:var(--text);cursor:pointer;
                         font-family:Cairo,sans-serif;font-weight:600;font-size:13px;
                         transition:all .2s;"
                  data-q="${qNum}" data-label="${label}">${label}</button>`
        ).join('')}
        <button onclick="setManualAnswer(${qNum},'SKIP',this)"
                style="padding:6px 10px;border-radius:6px;border:1px solid #ffb84f44;
                       background:#ffb84f11;color:var(--warn);cursor:pointer;
                       font-family:Cairo,sans-serif;font-weight:600;font-size:12px;"
                data-q="${qNum}" data-label="SKIP">تجاهل</button>
      </div>
    </div>`;
  }).join('');

  panel.style.display = 'block';
}

// ── Set manual answer for a question ──
function setManualAnswer(qNum, label, btn) {
  // Update visual selection
  const allBtns = btn.parentElement.querySelectorAll('button');
  allBtns.forEach(b => {
    b.style.background = 'var(--bg3)';
    b.style.color = 'var(--text)';
    b.style.borderColor = 'var(--border2)';
  });
  btn.style.background   = label === 'SKIP' ? '#ffb84f33' : 'var(--accent)';
  btn.style.color        = label === 'SKIP' ? 'var(--warn)' : '#fff';
  btn.style.borderColor  = label === 'SKIP' ? 'var(--warn)' : 'var(--accent)';

  // Store selection
  scannerState.currentScannedAnswers[qNum] = label === 'SKIP' ? null : label;
}

// ── Confirm manual answers and compute score ──
function confirmManualAnswers() {
  // Check all unclear questions were addressed
  const stillPending = scannerState.pendingManualQuestions.filter(qNum => {
    const ans = scannerState.currentScannedAnswers[qNum];
    return ans === null || ans === 'MULTI'; // null after SKIP is OK (null is stored)
  });

  if (stillPending.length > 0) {
    // Re-check: SKIP sets to null (allowed), MULTI is not allowed
    const reallyPending = scannerState.pendingManualQuestions.filter(qNum =>
      scannerState.currentScannedAnswers[qNum] === 'MULTI'
    );
    if (reallyPending.length > 0) {
      showToast(`يرجى تحديد إجابة للأسئلة: ${reallyPending.join(', ')}`, 'error');
      return;
    }
  }

  const model = scannerState.selectedModel ||
    scannerState.selectedSubject?.models[scannerState.selectedModelIdx || 0];
  if (!model) return;

  document.getElementById('manual-override-panel').style.display = 'none';
  computeAndShowScore(scannerState.currentScannedAnswers, model);
}

// ── Compute score and display ──
function computeAndShowScore(answers, model) {
  let correct = 0;
  const total = model.questions.length;

  model.questions.forEach((q, i) => {
    const qNum = i + 1;
    const studentAns = answers[qNum];
    if (studentAns && studentAns === q.correctLabel) correct++;
  });

  const scorePanel = document.getElementById('score-summary-panel');
  document.getElementById('score-display').textContent  = `${correct} / ${total}`;
  document.getElementById('score-percent').textContent  = `${Math.round((correct/total)*100)}%`;
  scorePanel.style.display = 'block';

  // Store result pending save
  scannerState._pendingResult = {
    studentId:  scannerState.currentStudentId,
    modelName:  model.name,
    modelLetter: model.letter,
    subject:    scannerState.selectedSubject?.name || '',
    answers:    { ...answers },
    correct,
    total,
    score:      correct,
    percent:    Math.round((correct/total)*100),
    timestamp:  new Date().toISOString()
  };
}

// ── Save result and prepare for next student ──
function saveAndNextStudent() {
  if (!scannerState._pendingResult) return;

  // Remove any existing entry for this student+subject
  scannerState.allResults = scannerState.allResults.filter(r =>
    !(r.studentId === scannerState._pendingResult.studentId &&
      r.subject   === scannerState._pendingResult.subject)
  );

  scannerState.allResults.push(scannerState._pendingResult);
  saveScanResults(scannerState.allResults);
  renderResultsTable();

  showToast(`تم حفظ نتيجة الطالب ${scannerState._pendingResult.studentId} ✓`, 'success');

  // Reset for next student
  scannerState._pendingResult = null;
  scannerState.currentStudentId = null;
  scannerState.currentScannedAnswers = null;
  scannerState.pendingManualQuestions = [];

  document.getElementById('scanner-current-student').textContent = '-';
  document.getElementById('scanner-student-id').value = '';
  document.getElementById('scanner-camera-card').style.display = 'none';
  document.getElementById('scan-result-content').textContent = 'اضغط على زر المسح لبدء العملية...';
  document.getElementById('manual-override-panel').style.display = 'none';
  document.getElementById('score-summary-panel').style.display = 'none';

  stopCamera();

  // Focus student ID for fast entry
  setTimeout(() => document.getElementById('scanner-student-id').focus(), 100);
}

// ── Rescan current student ──
function rescanCurrent() {
  scannerState.currentScannedAnswers = null;
  scannerState.pendingManualQuestions = [];
  scannerState._pendingResult = null;

  document.getElementById('scan-result-content').textContent = 'اضغط على زر المسح لبدء العملية...';
  document.getElementById('manual-override-panel').style.display = 'none';
  document.getElementById('score-summary-panel').style.display = 'none';

  if (!scannerState.cameraStream) startCamera();
}

// ── Render results table ──
function renderResultsTable() {
  const card   = document.getElementById('scanner-results-card');
  const tbody  = document.getElementById('results-tbody');
  const stats  = document.getElementById('results-stats');

  if (!tbody) return;

  // Filter by current subject if selected
  const subject = scannerState.selectedSubject?.name || '';
  const filtered = subject
    ? scannerState.allResults.filter(r => r.subject === subject)
    : scannerState.allResults;

  if (filtered.length === 0) {
    card.style.display = 'none';
    return;
  }

  card.style.display = 'block';

  tbody.innerHTML = filtered.map((r, i) => {
    const percent = r.percent ?? Math.round((r.correct/r.total)*100);
    const color   = percent >= 60 ? 'var(--accent3)' : percent >= 50 ? 'var(--warn)' : 'var(--danger)';
    return `<tr>
      <td>${i+1}</td>
      <td><strong style="color:var(--accent);">${r.studentId}</strong></td>
      <td>${r.modelName || r.modelLetter || '-'}</td>
      <td>${r.correct} / ${r.total}</td>
      <td><strong style="color:${color};">${r.score}</strong></td>
      <td><span style="color:${color};">${percent}%</span></td>
      <td>
        <button onclick="deleteResult('${r.studentId}','${r.subject}')"
                style="padding:4px 10px;background:#ff4f6a22;color:var(--danger);border:none;
                       border-radius:4px;cursor:pointer;font-family:Cairo,sans-serif;font-size:11px;">حذف</button>
      </td>
    </tr>`;
  }).join('');

  // Stats
  const avg = filtered.length > 0
    ? Math.round(filtered.reduce((s,r) => s + (r.percent??0), 0) / filtered.length)
    : 0;
  const passing = filtered.filter(r => (r.percent??0) >= 60).length;
  stats.innerHTML = `إجمالي الطلاب: <strong>${filtered.length}</strong> | متوسط الدرجات: <strong style="color:var(--accent);">${avg}%</strong> | الناجحون (≥60%): <strong style="color:var(--accent3);">${passing}</strong>`;
}

// ── Delete a result ──
function deleteResult(studentId, subject) {
  if (!confirm(`حذف نتيجة الطالب ${studentId}؟`)) return;
  scannerState.allResults = scannerState.allResults.filter(
    r => !(r.studentId === studentId && r.subject === subject)
  );
  saveScanResults(scannerState.allResults);
  renderResultsTable();
  showToast('تم حذف النتيجة', 'success');
}

// ── Clear all results ──
function clearAllResults() {
  if (!confirm('هل أنت متأكد من حذف جميع النتائج؟ لا يمكن التراجع عن هذا الإجراء.')) return;
  const subject = scannerState.selectedSubject?.name || '';
  if (subject) {
    scannerState.allResults = scannerState.allResults.filter(r => r.subject !== subject);
  } else {
    scannerState.allResults = [];
  }
  saveScanResults(scannerState.allResults);
  renderResultsTable();
  showToast('تم مسح النتائج', 'success');
}

// ── Export results to Excel ──
function exportResultsExcel() {
  const subject  = scannerState.selectedSubject?.name || '';
  const filtered = subject
    ? scannerState.allResults.filter(r => r.subject === subject)
    : scannerState.allResults;

  if (filtered.length === 0) { showToast('لا توجد نتائج للتصدير', 'error'); return; }

  // Build workbook
  const wb   = XLSX.utils.book_new();
  const rows = [
    ['الرقم الجامعي', 'النموذج', 'المادة', 'الإجابات الصحيحة', 'مجموع الأسئلة', 'العلامة', 'النسبة المئوية', 'تاريخ المسح']
  ];

  filtered.forEach(r => {
    const percent = r.percent ?? Math.round((r.correct/r.total)*100);
    rows.push([
      r.studentId,
      r.modelName || r.modelLetter || '',
      r.subject,
      r.correct,
      r.total,
      r.score,
      percent + '%',
      r.timestamp ? new Date(r.timestamp).toLocaleString('ar-EG') : ''
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);

  // Column widths
  ws['!cols'] = [
    {wch:18},{wch:14},{wch:20},{wch:16},{wch:14},{wch:10},{wch:14},{wch:22}
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'نتائج الامتحان');

  const filename = `نتائج_${subject || 'الامتحان'}_${new Date().toLocaleDateString('ar-EG').replace(/\//g,'-')}.xlsx`;
  XLSX.writeFile(wb, filename);
  showToast('تم تحميل ملف Excel ✓', 'success');
}

document.addEventListener('DOMContentLoaded', () => {
  initDropZone();
  initColorPickers();
});
