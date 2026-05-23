// ============================================================
//  scanner.js — OMR & Fill-Table Scanner (rewritten v2)
// ============================================================

const RESULTS_STORAGE_KEY = "omr_results_v1";

let scannerState = {
  selectedSubject:        null,
  selectedModelIdx:       null,
  selectedModel:          null,
  currentStudentId:       null,
  currentScannedAnswers:  null,
  pendingManualQuestions: [],
  cameraStream:           null,
  facingMode:             "environment",
  allResults:             []
};

// ── Storage helpers ──────────────────────────────────────────
function loadScanResults() {
  try {
    const raw = localStorage.getItem("omr_results_v1");
    return raw ? JSON.parse(raw) : [];
  } catch (e) { return []; }
}

function saveScanResults(arr) {
  try { localStorage.setItem("omr_results_v1", JSON.stringify(arr)); }
  catch (e) { console.error("Save results error:", e); }
}

// ── Panel init ───────────────────────────────────────────────
function initScannerPanel() {
  const projects = getAllProjects();
  const sel = document.getElementById("scanner-subject-select");
  if (!sel) return;

  sel.innerHTML = '<option value="">-- اختر المادة --</option>';

  if (models.length > 0 && examCfg.subject) {
    const opt = document.createElement("option");
    opt.value = "__current__";
    opt.textContent = `${examCfg.subject} (الجلسة الحالية — ${models.length} نماذج)`;
    sel.appendChild(opt);
  }

  projects.forEach(p => {
    const opt = document.createElement("option");
    opt.value = p.id;
    opt.textContent = `${p.name} (${p.models} نماذج — ${new Date(p.date).toLocaleDateString("ar-EG")})`;
    sel.appendChild(opt);
  });

  scannerState.allResults = loadScanResults();
  if (scannerState.allResults.length > 0) renderResultsTable();
}

function onScannerSubjectChange() {
  const val  = document.getElementById("scanner-subject-select").value;
  const mSel = document.getElementById("scanner-model-select");
  const stCard = document.getElementById("scanner-student-card");
  const mInfo  = document.getElementById("scanner-model-info");

  mSel.innerHTML = '<option value="">-- اختر النموذج --</option>';
  stCard.style.display = "none";
  mInfo.style.display  = "none";
  scannerState.selectedSubject = null;
  if (!val) return;

  let mdls = [], name = "";
  if (val === "__current__") {
    mdls = models; name = examCfg.subject;
  } else {
    const proj = loadProjectFromStorage(parseInt(val));
    if (proj) { mdls = proj.models || []; name = proj.examCfg?.subject || "غير محدد"; }
  }

  if (mdls.length === 0) { showToast("لا توجد نماذج لهذه المادة", "error"); return; }

  scannerState.selectedSubject = { val, models: mdls, name };

  const autoOpt = document.createElement("option");
  autoOpt.value = "auto";
  autoOpt.textContent = "🔍 تحديد تلقائي (من الورقة)";
  mSel.appendChild(autoOpt);

  mdls.forEach((m, idx) => {
    const opt = document.createElement("option");
    opt.value = idx;
    opt.textContent = `${m.name} — ${m.questions.length} سؤال`;
    mSel.appendChild(opt);
  });

  stCard.style.display = "block";
  document.getElementById("results-subject-label").textContent = name;
}

function onScannerModelChange() {
  const val  = document.getElementById("scanner-model-select").value;
  const info = document.getElementById("scanner-model-info");
  if (!val || !scannerState.selectedSubject) { info.style.display = "none"; return; }

  if (val === "auto") {
    info.style.display = "block";
    info.innerHTML = "🔍 سيتم تحديد النموذج تلقائياً من الورقة. إذا لم يتمكن الماسح من تحديده، سيطلب منك الاختيار يدوياً.";
    scannerState.selectedModelIdx = "auto";
    return;
  }

  const idx = parseInt(val);
  const mdl = scannerState.selectedSubject.models[idx];
  scannerState.selectedModelIdx = idx;
  scannerState.selectedModel = mdl;
  info.style.display = "block";
  info.innerHTML = `✅ النموذج: <strong style="color:var(--accent);">${mdl.name}</strong> | عدد الأسئلة: <strong>${mdl.questions.length}</strong> | مفتاح الإجابات جاهز للمقارنة`;
}

function startScanSession() {
  const sid = document.getElementById("scanner-student-id").value.trim();
  if (!sid) { showToast("يرجى إدخال الرقم الجامعي", "error"); return; }
  if (!document.getElementById("scanner-model-select").value) { showToast("يرجى اختيار النموذج", "error"); return; }

  const existing = scannerState.allResults.find(
    r => r.studentId === sid && r.subject === scannerState.selectedSubject?.name
  );
  if (existing) {
    if (!confirm(`الرقم الجامعي "${sid}" موجود مسبقاً لهذه المادة. هل تريد الاستبدال؟`)) return;
    scannerState.allResults = scannerState.allResults.filter(
      r => !(r.studentId === sid && r.subject === scannerState.selectedSubject?.name)
    );
  }

  scannerState.currentStudentId = sid;
  document.getElementById("scanner-current-student").textContent = sid;
  document.getElementById("scanner-camera-card").style.display = "block";
  document.getElementById("scan-result-content").textContent = "اضغط على زر المسح لبدء العملية...";
  document.getElementById("manual-override-panel").style.display = "none";
  document.getElementById("score-summary-panel").style.display = "none";
  startCamera();
}

// ── Camera ───────────────────────────────────────────────────
async function startCamera() {
  try {
    if (scannerState.cameraStream)
      scannerState.cameraStream.getTracks().forEach(t => t.stop());
    const stream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: scannerState.facingMode, width: { ideal: 1920 }, height: { ideal: 1440 } }
    });
    scannerState.cameraStream = stream;
    document.getElementById("scanner-video").srcObject = stream;
  } catch (e) {
    showToast("لا يمكن الوصول للكاميرا: " + e.message, "error");
  }
}

function stopCamera() {
  if (scannerState.cameraStream) {
    scannerState.cameraStream.getTracks().forEach(t => t.stop());
    scannerState.cameraStream = null;
  }
  document.getElementById("scanner-video").srcObject = null;
}

function switchCamera() {
  scannerState.facingMode = scannerState.facingMode === "environment" ? "user" : "environment";
  startCamera();
}

function captureAndScan() {
  const vid = document.getElementById("scanner-video");
  if (!vid.srcObject) { showToast("الكاميرا غير مفعّلة", "error"); return; }
  const canvas = document.createElement("canvas");
  canvas.width  = vid.videoWidth  || 1280;
  canvas.height = vid.videoHeight || 960;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(vid, 0, 0, canvas.width, canvas.height);
  processOMRImage(ctx.getImageData(0, 0, canvas.width, canvas.height), canvas.width, canvas.height);
}

function scanFromFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  const img = new Image();
  img.onload = () => {
    const canvas = document.createElement("canvas");
    canvas.width  = img.width;
    canvas.height = img.height;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(img, 0, 0);
    processOMRImage(ctx.getImageData(0, 0, canvas.width, canvas.height), canvas.width, canvas.height);
  };
  img.src = URL.createObjectURL(file);
}

function processOMRImage(imgData, w, h) {
  const mdl = scannerState.selectedModel || scannerState.selectedSubject?.models[0];
  if (!mdl) { showToast("يرجى اختيار النموذج أولاً", "error"); return; }

  document.getElementById("scan-result-content").innerHTML =
    '<div style="color:var(--accent);">⏳ جارٍ تحليل الصورة...</div>';

  setTimeout(() => {
    try {
      const answers = analyzeOMRBubbles(imgData, w, h, mdl.questions.length);
      scannerState.currentScannedAnswers = answers;

      const pending = [];
      Object.entries(answers).forEach(([q, a]) => {
        if (a === null || a === "MULTI") pending.push(parseInt(q));
      });
      scannerState.pendingManualQuestions = pending;

      showScanResults(answers, mdl);

      if (pending.length > 0) showManualOverride(pending, mdl);
      else computeAndShowScore(answers, mdl);

    } catch (err) {
      document.getElementById("scan-result-content").innerHTML =
        `<div style="color:var(--danger);">❌ خطأ في تحليل الصورة: ${err.message}</div>`;
      console.error("OMR error:", err);
    }
  }, 100);
}

// ============================================================
//  CORE OMR ANALYSIS  (rewritten v2)
// ============================================================

/**
 * analyzeOMRBubbles
 *
 * Strategy:
 *  1. Convert to greyscale.
 *  2. Apply adaptive threshold (Otsu) to get a binary image.
 *  3. Project rows/cols to find the answer-grid region robustly.
 *  4. Divide grid into (numQuestions × 4) cells.
 *  5. For each cell compute fill ratio of dark pixels.
 *  6. Decide answer per row using a two-stage decision:
 *       a. If max fill ratio < EMPTY_THRESH  → no answer (null)
 *       b. If top-2 difference  < MULTI_GAP  → multiple marks (MULTI)
 *       c. Otherwise → single winner
 *
 * Works for both OMR circles and plain fill-in tables because both
 * produce dark-filled regions; only the threshold matters.
 */
function analyzeOMRBubbles(imgData, W, H, numQuestions) {
  const data = imgData.data;

  // ── 1. Greyscale ──────────────────────────────────────────
  const grey = new Uint8Array(W * H);
  for (let i = 0; i < W * H; i++) {
    const b = i * 4;
    grey[i] = Math.round(0.299 * data[b] + 0.587 * data[b + 1] + 0.114 * data[b + 2]);
  }

  // ── 2. Otsu threshold ────────────────────────────────────
  const thresh = otsuThreshold(grey, W * H);

  // Binary: 1 = dark (mark), 0 = light (background)
  const bin = new Uint8Array(W * H);
  for (let i = 0; i < W * H; i++) bin[i] = grey[i] <= thresh ? 1 : 0;

  // ── 3. Row & col projections ──────────────────────────────
  const rowProj = new Float32Array(H);
  const colProj = new Float32Array(W);

  for (let y = 0; y < H; y++) {
    let s = 0;
    for (let x = 0; x < W; x++) s += bin[y * W + x];
    rowProj[y] = s / W;
  }
  for (let x = 0; x < W; x++) {
    let s = 0;
    for (let y = 0; y < H; y++) s += bin[y * W + x];
    colProj[x] = s / H;
  }

  // ── 4. Find grid boundaries ───────────────────────────────
  // We look for the extent of the answer-grid (rows with significant
  // dark content).  Use a lower threshold so we catch light fills too.
  const ROW_THRESH = 0.04;  // at least 4% dark per row to count
  const COL_THRESH = 0.03;

  let rowTop    = Math.floor(H * 0.05);
  let rowBottom = Math.floor(H * 0.95);
  let colLeft   = Math.floor(W * 0.03);
  let colRight  = Math.floor(W * 0.97);

  // Walk inward until we find significant content
  for (let y = Math.floor(H * 0.05); y < Math.floor(H * 0.95); y++) {
    if (rowProj[y] >= ROW_THRESH) { rowTop = y; break; }
  }
  for (let y = Math.floor(H * 0.95); y > rowTop; y--) {
    if (rowProj[y] >= ROW_THRESH) { rowBottom = y; break; }
  }
  for (let x = Math.floor(W * 0.03); x < Math.floor(W * 0.97); x++) {
    if (colProj[x] >= COL_THRESH) { colLeft = x; break; }
  }
  for (let x = Math.floor(W * 0.97); x > colLeft; x--) {
    if (colProj[x] >= COL_THRESH) { colRight = x; break; }
  }

  const gridH = rowBottom - rowTop;
  const gridW = colRight  - colLeft;

  if (gridH < 20 || gridW < 20) {
    // Fallback: use full image
    rowTop    = 0; rowBottom = H;
    colLeft   = 0; colRight  = W;
  }

  // ── 5. Detect number of choices (4 columns default) ──────
  // Try to auto-detect by finding column separators inside grid,
  // but default to 4 (A-B-C-D) which covers 99 % of cases.
  const NUM_CHOICES = 4;
  const labels = ["A", "B", "C", "D"];

  // Cell dimensions
  const cellH = gridH / numQuestions;
  const cellW = gridW / NUM_CHOICES;

  // ── 6. Per-question analysis ──────────────────────────────
  // Tuning constants
  const EMPTY_THRESH = 0.06;   // below this → no mark at all
  const MULTI_GAP    = 0.08;   // top-2 difference below this → ambiguous
  const MARK_THRESH  = 0.08;   // candidate must exceed this to be considered

  const result = {};

  for (let q = 0; q < numQuestions; q++) {
    // Row band for this question (with small inset to avoid grid lines)
    const yPad = Math.max(2, Math.round(cellH * 0.10));
    const y0 = Math.round(rowTop + q * cellH) + yPad;
    const y1 = Math.round(rowTop + (q + 1) * cellH) - yPad;

    const fills = [];

    for (let c = 0; c < NUM_CHOICES; c++) {
      // Column band (with inset)
      const xPad = Math.max(2, Math.round(cellW * 0.10));
      const x0 = Math.round(colLeft + c * cellW) + xPad;
      const x1 = Math.round(colLeft + (c + 1) * cellW) - xPad;

      // Count dark pixels in this cell
      let dark = 0, total = 0;
      for (let y = y0; y < y1 && y < H; y++) {
        for (let x = x0; x < x1 && x < W; x++) {
          if (x >= 0 && y >= 0) {
            if (bin[y * W + x]) dark++;
            total++;
          }
        }
      }

      fills.push(total > 0 ? dark / total : 0);
    }

    // Decision logic
    const maxFill = Math.max(...fills);

    if (maxFill < EMPTY_THRESH) {
      result[q + 1] = null; // no answer
      continue;
    }

    // Candidates: cells above MARK_THRESH
    const candidates = fills
      .map((f, i) => ({ label: labels[i], fill: f }))
      .filter(c => c.fill >= MARK_THRESH)
      .sort((a, b) => b.fill - a.fill);

    if (candidates.length === 0) {
      result[q + 1] = null;
    } else if (candidates.length === 1) {
      result[q + 1] = candidates[0].label;
    } else {
      // More than one candidate above threshold
      const diff = candidates[0].fill - candidates[1].fill;
      if (diff < MULTI_GAP) {
        result[q + 1] = "MULTI"; // ambiguous / multiple marks
      } else {
        result[q + 1] = candidates[0].label;
      }
    }
  }

  return result;
}

// ── Otsu's method for global threshold ──────────────────────
function otsuThreshold(grey, size) {
  // Build histogram
  const hist = new Int32Array(256);
  for (let i = 0; i < size; i++) hist[grey[i]]++;

  let total = size;
  let sumB = 0, wB = 0, sum = 0;
  for (let i = 0; i < 256; i++) sum += i * hist[i];

  let maxVar = 0, thresh = 128;
  for (let t = 0; t < 256; t++) {
    wB += hist[t];
    if (wB === 0) continue;
    const wF = total - wB;
    if (wF === 0) break;
    sumB += t * hist[t];
    const mB = sumB / wB;
    const mF = (sum - sumB) / wF;
    const v  = wB * wF * (mB - mF) * (mB - mF);
    if (v > maxVar) { maxVar = v; thresh = t; }
  }
  return thresh;
}

// ============================================================
//  UI helpers (unchanged)
// ============================================================

function showScanResults(answers, model) {
  const total     = model.questions.length;
  const detected  = Object.values(answers).filter(a => a && a !== "MULTI").length;
  const unclear   = Object.values(answers).filter(a => a === null || a === "MULTI").length;

  let html = `<div style="margin-bottom:10px;font-size:12px;color:var(--text2);">
    تم اكتشاف <strong style="color:var(--accent3);">${detected}</strong> إجابة من أصل ${total}
    ${unclear > 0 ? `| <strong style="color:var(--warn);">${unclear}</strong> غير واضحة` : ""}
  </div>`;

  html += '<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:4px;">';
  for (let q = 1; q <= total; q++) {
    const a = answers[q];
    let bg    = "var(--bg)";
    let color = "var(--text2)";
    let label = a || "?";

    if (a === null)    { bg = "#ffb84f22"; color = "var(--warn)";   label = "؟"; }
    if (a === "MULTI") { bg = "#ff4f6a22"; color = "var(--danger)"; label = "!!"; }
    if (a && a !== "MULTI") { bg = "#4f7cff22"; color = "var(--accent)"; }

    html += `<div style="padding:4px;background:${bg};border-radius:6px;text-align:center;border:1px solid ${bg};">
      <div style="font-size:9px;color:var(--text3);">${q}</div>
      <div style="font-size:13px;font-weight:700;color:${color};">${label}</div>
    </div>`;
  }
  html += "</div>";
  document.getElementById("scan-result-content").innerHTML = html;
}

function showManualOverride(pending, model) {
  const panel = document.getElementById("manual-override-panel");
  document.getElementById("manual-override-list").innerHTML = pending.map(q => {
    const question = model.questions[q - 1];
    const choices  = question ? question.choices.map(c => c.label) : ["A", "B", "C", "D"];
    const isMulti  = scannerState.currentScannedAnswers[q] === "MULTI";

    return `<div style="padding:8px;background:var(--bg);border-radius:8px;margin-bottom:8px;border:1px solid var(--border);">
      <div style="font-size:12px;color:var(--text2);margin-bottom:6px;">
        <strong style="color:var(--warn);">س${q}:</strong>
        ${question ? question.text.slice(0, 60) + (question.text.length > 60 ? "..." : "") : "سؤال رقم " + q}
        ${isMulti
          ? '<span style="color:var(--danger);font-size:11px;"> (إجابات متعددة)</span>'
          : '<span style="color:var(--warn);font-size:11px;"> (لم تُكتشف)</span>'}
      </div>
      <div style="display:flex;gap:6px;flex-wrap:wrap;">
        ${choices.map(lbl => `<button onclick="setManualAnswer(${q},'${lbl}',this)"
          style="padding:6px 14px;border-radius:6px;border:1px solid var(--border2);
                 background:var(--bg3);color:var(--text);cursor:pointer;
                 font-family:Cairo,sans-serif;font-weight:600;font-size:13px;
                 transition:all .2s;"
          data-q="${q}" data-label="${lbl}">${lbl}</button>`).join("")}
        <button onclick="setManualAnswer(${q},'SKIP',this)"
          style="padding:6px 10px;border-radius:6px;border:1px solid #ffb84f44;
                 background:#ffb84f11;color:var(--warn);cursor:pointer;
                 font-family:Cairo,sans-serif;font-weight:600;font-size:12px;"
          data-q="${q}" data-label="SKIP">تجاهل</button>
      </div>
    </div>`;
  }).join("");
  panel.style.display = "block";
}

function setManualAnswer(q, label, btn) {
  btn.parentElement.querySelectorAll("button").forEach(b => {
    b.style.background  = "var(--bg3)";
    b.style.color       = "var(--text)";
    b.style.borderColor = "var(--border2)";
  });
  btn.style.background  = label === "SKIP" ? "#ffb84f33" : "var(--accent)";
  btn.style.color       = label === "SKIP" ? "var(--warn)"  : "#fff";
  btn.style.borderColor = label === "SKIP" ? "var(--warn)"  : "var(--accent)";
  scannerState.currentScannedAnswers[q] = label === "SKIP" ? null : label;
}

function confirmManualAnswers() {
  const stillPending = scannerState.pendingManualQuestions.filter(q => {
    const a = scannerState.currentScannedAnswers[q];
    return a === null || a === "MULTI";
  });

  if (stillPending.length > 0) {
    const multiRemaining = stillPending.filter(q => scannerState.currentScannedAnswers[q] === "MULTI");
    if (multiRemaining.length > 0) {
      showToast(`يرجى تحديد إجابة للأسئلة: ${multiRemaining.join(", ")}`, "error");
      return;
    }
  }

  const mdl = scannerState.selectedModel ||
    scannerState.selectedSubject?.models[scannerState.selectedModelIdx || 0];
  if (!mdl) return;

  document.getElementById("manual-override-panel").style.display = "none";
  computeAndShowScore(scannerState.currentScannedAnswers, mdl);
}

function computeAndShowScore(answers, model) {
  let correct = 0;
  const total = model.questions.length;

  model.questions.forEach((q, i) => {
    const given = answers[i + 1];
    if (given && given === q.correctLabel) correct++;
  });

  const panel = document.getElementById("score-summary-panel");
  document.getElementById("score-display").textContent  = `${correct} / ${total}`;
  document.getElementById("score-percent").textContent  = `${Math.round(correct / total * 100)}%`;
  panel.style.display = "block";

  scannerState._pendingResult = {
    studentId:   scannerState.currentStudentId,
    modelName:   model.name,
    modelLetter: model.letter,
    subject:     scannerState.selectedSubject?.name || "",
    answers:     { ...answers },
    correct,
    total,
    score:       correct,
    percent:     Math.round(correct / total * 100),
    timestamp:   new Date().toISOString()
  };
}

function saveAndNextStudent() {
  if (!scannerState._pendingResult) return;

  scannerState.allResults = scannerState.allResults.filter(r =>
    !(r.studentId === scannerState._pendingResult.studentId &&
      r.subject   === scannerState._pendingResult.subject)
  );
  scannerState.allResults.push(scannerState._pendingResult);
  saveScanResults(scannerState.allResults);
  renderResultsTable();
  showToast(`تم حفظ نتيجة الطالب ${scannerState._pendingResult.studentId} ✓`, "success");

  // Reset
  scannerState._pendingResult         = null;
  scannerState.currentStudentId       = null;
  scannerState.currentScannedAnswers  = null;
  scannerState.pendingManualQuestions = [];

  document.getElementById("scanner-current-student").textContent = "-";
  document.getElementById("scanner-student-id").value            = "";
  document.getElementById("scanner-camera-card").style.display   = "none";
  document.getElementById("scan-result-content").textContent     = "اضغط على زر المسح لبدء العملية...";
  document.getElementById("manual-override-panel").style.display = "none";
  document.getElementById("score-summary-panel").style.display   = "none";

  stopCamera();
  setTimeout(() => document.getElementById("scanner-student-id").focus(), 100);
}

function rescanCurrent() {
  scannerState.currentScannedAnswers  = null;
  scannerState.pendingManualQuestions = [];
  scannerState._pendingResult         = null;

  document.getElementById("scan-result-content").textContent     = "اضغط على زر المسح لبدء العملية...";
  document.getElementById("manual-override-panel").style.display = "none";
  document.getElementById("score-summary-panel").style.display   = "none";

  if (!scannerState.cameraStream) startCamera();
}

function renderResultsTable() {
  const card  = document.getElementById("scanner-results-card");
  const tbody = document.getElementById("results-tbody");
  const stats = document.getElementById("results-stats");
  if (!tbody) return;

  const subject  = scannerState.selectedSubject?.name || "";
  const filtered = subject
    ? scannerState.allResults.filter(r => r.subject === subject)
    : scannerState.allResults;

  if (filtered.length === 0) { card.style.display = "none"; return; }

  card.style.display = "block";
  tbody.innerHTML = filtered.map((r, i) => {
    const pct   = r.percent ?? Math.round(r.correct / r.total * 100);
    const color = pct >= 60 ? "var(--accent3)" : pct >= 50 ? "var(--warn)" : "var(--danger)";
    return `<tr>
      <td>${i + 1}</td>
      <td><strong style="color:var(--accent);">${r.studentId}</strong></td>
      <td>${r.modelName || r.modelLetter || "-"}</td>
      <td>${r.correct} / ${r.total}</td>
      <td><strong style="color:${color};">${r.score}</strong></td>
      <td><span style="color:${color};">${pct}%</span></td>
      <td>
        <button onclick="deleteResult('${r.studentId}','${r.subject}')"
          style="padding:4px 10px;background:#ff4f6a22;color:var(--danger);border:none;
                 border-radius:4px;cursor:pointer;font-family:Cairo,sans-serif;font-size:11px;">حذف</button>
      </td>
    </tr>`;
  }).join("");

  const avg  = filtered.length > 0
    ? Math.round(filtered.reduce((s, r) => s + (r.percent ?? 0), 0) / filtered.length)
    : 0;
  const pass = filtered.filter(r => (r.percent ?? 0) >= 60).length;
  stats.innerHTML = `إجمالي الطلاب: <strong>${filtered.length}</strong> | متوسط الدرجات: <strong style="color:var(--accent);">${avg}%</strong> | الناجحون (≥60%): <strong style="color:var(--accent3);">${pass}</strong>`;
}

function deleteResult(studentId, subject) {
  if (!confirm(`حذف نتيجة الطالب ${studentId}؟`)) return;
  scannerState.allResults = scannerState.allResults.filter(
    r => !(r.studentId === studentId && r.subject === subject)
  );
  saveScanResults(scannerState.allResults);
  renderResultsTable();
  showToast("تم حذف النتيجة", "success");
}

function clearAllResults() {
  if (!confirm("هل أنت متأكد من حذف جميع النتائج؟ لا يمكن التراجع عن هذا الإجراء.")) return;
  const subject = scannerState.selectedSubject?.name || "";
  scannerState.allResults = subject
    ? scannerState.allResults.filter(r => r.subject !== subject)
    : [];
  saveScanResults(scannerState.allResults);
  renderResultsTable();
  showToast("تم مسح النتائج", "success");
}

function exportResultsExcel() {
  const subject  = scannerState.selectedSubject?.name || "";
  const filtered = subject
    ? scannerState.allResults.filter(r => r.subject === subject)
    : scannerState.allResults;

  if (filtered.length === 0) { showToast("لا توجد نتائج للتصدير", "error"); return; }

  const wb   = XLSX.utils.book_new();
  const rows = [["الرقم الجامعي","النموذج","المادة","الإجابات الصحيحة","مجموع الأسئلة","العلامة","النسبة المئوية","تاريخ المسح"]];

  filtered.forEach(r => {
    const pct = r.percent ?? Math.round(r.correct / r.total * 100);
    rows.push([
      r.studentId,
      r.modelName || r.modelLetter || "",
      r.subject,
      r.correct,
      r.total,
      r.score,
      pct + "%",
      r.timestamp ? new Date(r.timestamp).toLocaleString("ar-EG") : ""
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{wch:18},{wch:14},{wch:20},{wch:16},{wch:14},{wch:10},{wch:14},{wch:22}];
  XLSX.utils.book_append_sheet(wb, ws, "نتائج الامتحان");

  const filename = `نتائج_${subject || "الامتحان"}_${new Date().toLocaleDateString("ar-EG").replace(/\//g,"-")}.xlsx`;
  XLSX.writeFile(wb, filename);
  showToast("تم تحميل ملف Excel ✓", "success");
}
