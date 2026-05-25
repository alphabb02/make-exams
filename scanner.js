const RESULTS_STORAGE_KEY="omr_results_v1";let scannerState={selectedSubject:null,selectedModelIdx:null,selectedModel:null,currentStudentId:null,currentScannedAnswers:null,pendingManualQuestions:[],cameraStream:null,facingMode:"environment",allResults:[]};function loadScanResults(){try{const e=localStorage.getItem("omr_results_v1");return e?JSON.parse(e):[]}catch(e){return[]}}function saveScanResults(e){try{localStorage.setItem("omr_results_v1",JSON.stringify(e))}catch(e){console.error("Save results error:",e)}}function initScannerPanel(){const e=getAllProjects(),t=document.getElementById("scanner-subject-select");if(t){if(t.innerHTML='<option value="">-- اختر المادة --</option>',models.length>0&&examCfg.subject){const e=document.createElement("option");e.value="__current__",e.textContent=`${examCfg.subject} (الجلسة الحالية — ${models.length} نماذج)`,t.appendChild(e)}e.forEach(e=>{const n=document.createElement("option");n.value=e.id,n.textContent=`${e.name} (${e.models} نماذج — ${new Date(e.date).toLocaleDateString("ar-EG")})`,t.appendChild(n)}),scannerState.allResults=loadScanResults(),scannerState.allResults.length>0&&renderResultsTable()}}function onScannerSubjectChange(){const e=document.getElementById("scanner-subject-select").value,t=document.getElementById("scanner-model-select"),n=document.getElementById("scanner-student-card"),a=document.getElementById("scanner-model-info");if(t.innerHTML='<option value="">-- اختر النموذج --</option>',n.style.display="none",a.style.display="none",scannerState.selectedSubject=null,!e)return;let r=[],s="";if("__current__"===e)r=models,s=examCfg.subject;else{const t=loadProjectFromStorage(parseInt(e));t&&(r=t.models||[],s=t.examCfg?.subject||"غير محدد")}if(0===r.length)return void showToast("لا توجد نماذج لهذه المادة","error");scannerState.selectedSubject={val:e,models:r,name:s},r.forEach((e,n)=>{const a=document.createElement("option");a.value=n,a.textContent=`${e.name} — ${e.questions.length} سؤال`,t.appendChild(a)});const o=document.createElement("option");o.value="auto",o.textContent="🔍 تحديد تلقائي (من الورقة)",t.insertBefore(o,t.children[1]),n.style.display="block",document.getElementById("results-subject-label").textContent=s}function onScannerModelChange(){const e=document.getElementById("scanner-model-select").value,t=document.getElementById("scanner-model-info");if(!e||!scannerState.selectedSubject)return void(t.style.display="none");if("auto"===e)return t.style.display="block",t.innerHTML="🔍 سيتم تحديد النموذج تلقائياً من الورقة. إذا لم يتمكن الماسح من تحديده، سيطلب منك الاختيار يدوياً.",void(scannerState.selectedModelIdx="auto");const n=parseInt(e),a=scannerState.selectedSubject.models[n];scannerState.selectedModelIdx=n,scannerState.selectedModel=a,t.style.display="block",t.innerHTML=`✅ النموذج: <strong style="color:var(--accent);">${a.name}</strong> | عدد الأسئلة: <strong>${a.questions.length}</strong> | مفتاح الإجابات جاهز للمقارنة`}function startScanSession(){const e=document.getElementById("scanner-student-id").value.trim();if(e)if(document.getElementById("scanner-model-select").value){if(scannerState.allResults.find(t=>t.studentId===e&&t.subject===scannerState.selectedSubject?.name)){if(!confirm(`الرقم الجامعي "${e}" موجود مسبقاً لهذه المادة. هل تريد الاستبدال؟`))return;scannerState.allResults=scannerState.allResults.filter(t=>!(t.studentId===e&&t.subject===scannerState.selectedSubject?.name))}scannerState.currentStudentId=e,document.getElementById("scanner-current-student").textContent=e,document.getElementById("scanner-camera-card").style.display="block",document.getElementById("scan-result-content").textContent="اضغط على زر المسح لبدء العملية...",document.getElementById("manual-override-panel").style.display="none",document.getElementById("score-summary-panel").style.display="none",startCamera()}else showToast("يرجى اختيار النموذج","error");else showToast("يرجى إدخال الرقم الجامعي","error")}async function startCamera(){try{scannerState.cameraStream&&scannerState.cameraStream.getTracks().forEach(e=>e.stop());const e=await navigator.mediaDevices.getUserMedia({video:{facingMode:scannerState.facingMode,width:{ideal:1280},height:{ideal:960}}});scannerState.cameraStream=e,document.getElementById("scanner-video").srcObject=e}catch(e){showToast("لا يمكن الوصول للكاميرا: "+e.message,"error")}}function stopCamera(){scannerState.cameraStream&&(scannerState.cameraStream.getTracks().forEach(e=>e.stop()),scannerState.cameraStream=null),document.getElementById("scanner-video").srcObject=null}function switchCamera(){scannerState.facingMode="environment"===scannerState.facingMode?"user":"environment",startCamera()}function captureAndScan(){const e=document.getElementById("scanner-video");if(!e.srcObject)return void showToast("الكاميرا غير مفعّلة","error");const t=document.createElement("canvas");t.width=e.videoWidth||640,t.height=e.videoHeight||480;const n=t.getContext("2d");n.drawImage(e,0,0,t.width,t.height),processOMRImage(n.getImageData(0,0,t.width,t.height),t.width,t.height)}function scanFromFile(e){const t=e.target.files[0];if(!t)return;const n=new Image;n.onload=()=>{const e=document.createElement("canvas");e.width=n.width,e.height=n.height;const t=e.getContext("2d");t.drawImage(n,0,0),processOMRImage(t.getImageData(0,0,e.width,e.height),e.width,e.height)},n.src=URL.createObjectURL(t)}// ============================================================
//  AI-POWERED OMR ENGINE  —  analyzeOMRBubbles + processOMRImage
//  Uses Claude Vision API as primary engine,
//  with a robust pixel-based fallback.
// ============================================================

// ── Helper: convert ImageData → base64 JPEG via offscreen canvas ──
function imageDataToBase64(imageData, W, H) {
  const canvas = document.createElement('canvas');
  canvas.width = W; canvas.height = H;
  canvas.getContext('2d').putImageData(imageData, 0, 0);
  // compress to JPEG 85% to keep payload small
  return canvas.toDataURL('image/jpeg', 0.85).split(',')[1];
}

// ── Build the AI prompt ──
function buildOMRPrompt(numQuestions) {
  return `You are an expert OMR (Optical Mark Recognition) system analyzing an exam answer sheet image.

The answer sheet contains a GRID TABLE where:
- ROWS = answer choices (A, B, C, D) — there are 4 rows of bubbles
- COLUMNS = question numbers (1 to ${numQuestions})
- Each cell contains a circle/bubble
- A FILLED (darkened/shaded) bubble = the student's chosen answer
- An EMPTY bubble = not chosen

Your task: For EACH question number from 1 to ${numQuestions}, identify which bubble is filled.

IMPORTANT RULES:
1. Look carefully at the bubble fill level — filled bubbles are clearly darker than empty ones
2. If NO bubble is filled for a question → return null
3. If MULTIPLE bubbles appear filled for a question → return "MULTI"
4. The table may have a header row and a label column — ignore those
5. There may be TWO grid sections if questions exceed 20 (e.g. Q1-20 then Q21-22)
6. Be precise — even partial fills count if they're clearly darker than others in the same column

Respond with ONLY a valid JSON object, nothing else. Format:
{
  "1": "A",
  "2": "C", 
  "3": null,
  "4": "MULTI",
  ...
  "${numQuestions}": "B"
}

Keys must be string numbers "1" through "${numQuestions}". Values: "A", "B", "C", "D", null, or "MULTI".`;
}

// ── AI-based analysis using Claude Vision ──
async function analyzeWithAI(imageData, W, H, numQuestions) {
  const base64Image = imageDataToBase64(imageData, W, H);

  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1000,
      messages: [{
        role: 'user',
        content: [
          {
            type: 'image',
            source: { type: 'base64', media_type: 'image/jpeg', data: base64Image }
          },
          {
            type: 'text',
            text: buildOMRPrompt(numQuestions)
          }
        ]
      }]
    })
  });

  if (!response.ok) throw new Error(`API error: ${response.status}`);
  const data = await response.json();
  const text = data.content.map(b => b.text || '').join('');

  // Parse JSON from response
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) throw new Error('No JSON found in AI response');
  const raw = JSON.parse(jsonMatch[0]);

  // Normalize: ensure all questions 1..N are present
  const results = {};
  for (let q = 1; q <= numQuestions; q++) {
    const val = raw[String(q)];
    if (val === null || val === undefined) results[q] = null;
    else if (['A','B','C','D','MULTI'].includes(String(val).toUpperCase())) {
      results[q] = String(val).toUpperCase() === 'MULTI' ? 'MULTI' : String(val).toUpperCase();
    } else {
      results[q] = null;
    }
  }
  return results;
}

// ── Pixel-based fallback engine (improved) ──
function analyzeWithPixels(imageData, W, H, numQuestions) {
  const data = imageData.data;

  // 1. Grayscale
  const gray = new Uint8Array(W * H);
  for (let i = 0; i < W * H; i++) {
    gray[i] = Math.round(0.299 * data[4*i] + 0.587 * data[4*i+1] + 0.114 * data[4*i+2]);
  }

  // 2. Otsu threshold
  const hist = new Array(256).fill(0);
  for (let i = 0; i < gray.length; i++) hist[gray[i]]++;
  let sumAll = 0, sumB = 0, wB = 0, maxVar = 0, threshold = 128;
  for (let i = 0; i < 256; i++) sumAll += i * hist[i];
  for (let t = 0; t < 256; t++) {
    wB += hist[t]; if (!wB) continue;
    const wF = gray.length - wB; if (!wF) break;
    sumB += t * hist[t];
    const mB = sumB / wB, mF = (sumAll - sumB) / wF;
    const v = wB * wF * (mB - mF) ** 2;
    if (v > maxVar) { maxVar = v; threshold = t; }
  }
  const binary = new Uint8Array(W * H);
  for (let i = 0; i < gray.length; i++) binary[i] = gray[i] < threshold ? 1 : 0;

  // 3. Row/Col density
  const rowDark = new Float32Array(H);
  const colDark = new Float32Array(W);
  for (let y = 0; y < H; y++) {
    let cnt = 0;
    for (let x = 0; x < W; x++) if (binary[y * W + x]) cnt++;
    rowDark[y] = cnt / W;
  }
  for (let x = 0; x < W; x++) {
    let cnt = 0;
    for (let y = 0; y < H; y++) if (binary[y * W + x]) cnt++;
    colDark[x] = cnt / H;
  }

  // 4. Detect grid lines
  function findLines(profile, size, minDensity, minGap) {
    const lines = [];
    let inLine = false, ls = 0;
    for (let i = 0; i < size; i++) {
      if (profile[i] >= minDensity) {
        if (!inLine) { inLine = true; ls = i; }
      } else if (inLine) {
        inLine = false;
        const mid = Math.floor((ls + i) / 2);
        if (!lines.length || mid - lines[lines.length-1] >= minGap) lines.push(mid);
      }
    }
    if (inLine) {
      const mid = Math.floor((ls + size) / 2);
      if (!lines.length || mid - lines[lines.length-1] >= minGap) lines.push(mid);
    }
    return lines;
  }

  const hLines = findLines(rowDark, H, 0.12, 6);
  const vLines = findLines(colDark, W, 0.06, 5);

  // 5. Build row bounds (4 choice rows: A,B,C,D)
  const NUM_CHOICES = 4;
  let rowBounds;
  if (hLines.length >= NUM_CHOICES + 1) {
    rowBounds = hLines.slice(hLines.length - (NUM_CHOICES + 1));
  } else if (hLines.length >= 2) {
    const top = hLines[0], bot = hLines[hLines.length-1];
    rowBounds = Array.from({length: NUM_CHOICES + 1}, (_, i) =>
      Math.floor(top + (bot - top) * i / NUM_CHOICES));
  } else {
    const top = Math.floor(H * 0.3), bot = Math.floor(H * 0.85);
    rowBounds = Array.from({length: NUM_CHOICES + 1}, (_, i) =>
      Math.floor(top + (bot - top) * i / NUM_CHOICES));
  }

  // 6. Build column bounds (numQuestions question columns, skip label col)
  let colBounds;
  if (vLines.length >= numQuestions + 1) {
    colBounds = vLines.slice(vLines.length - (numQuestions + 1));
    colBounds.push(Math.floor(W * 0.99));
  } else if (vLines.length >= 2) {
    const labelEnd = vLines[0];
    const right = vLines[vLines.length-1];
    colBounds = [labelEnd];
    for (let i = 0; i <= numQuestions; i++)
      colBounds.push(Math.floor(labelEnd + (right - labelEnd) * i / numQuestions));
  } else {
    const labelEnd = Math.floor(W * 0.12);
    const right = Math.floor(W * 0.99);
    colBounds = [labelEnd];
    for (let i = 0; i <= numQuestions; i++)
      colBounds.push(Math.floor(labelEnd + (right - labelEnd) * i / numQuestions));
  }

  // 7. Score each cell
  const LABELS = ['A','B','C','D'];
  const MARGIN = 0.18;
  const results = {};

  for (let q = 0; q < numQuestions; q++) {
    const cx0 = colBounds[q], cx1 = colBounds[q + 1];
    if (!cx1 || cx1 <= cx0) { results[q+1] = null; continue; }
    const cw = cx1 - cx0;
    const mx = Math.max(2, Math.floor(cw * MARGIN));
    const ratios = [];

    for (let c = 0; c < NUM_CHOICES; c++) {
      const ry0 = rowBounds[c], ry1 = rowBounds[c+1];
      if (!ry1 || ry1 <= ry0) { ratios.push(0); continue; }
      const rh = ry1 - ry0;
      const my = Math.max(2, Math.floor(rh * MARGIN));
      let dark = 0, total = 0;
      for (let y = ry0 + my; y < ry1 - my; y++) {
        for (let x = cx0 + mx; x < cx1 - mx; x++) {
          if (x >= 0 && x < W && y >= 0 && y < H) {
            if (binary[y * W + x]) dark++;
            total++;
          }
        }
      }
      ratios.push(total > 0 ? dark / total : 0);
    }

    const maxR = Math.max(...ratios);
    const ABS = 0.10, REL = 0.50;
    const cands = ratios
      .map((r, i) => ({ label: LABELS[i], ratio: r }))
      .filter(x => x.ratio >= ABS && x.ratio >= maxR * REL);

    if (!cands.length) results[q+1] = null;
    else if (cands.length === 1) results[q+1] = cands[0].label;
    else {
      cands.sort((a,b) => b.ratio - a.ratio);
      const gap = cands[0].ratio - cands[1].ratio;
      results[q+1] = gap < cands[0].ratio * 0.25 ? 'MULTI' : cands[0].label;
    }
  }
  return results;
}

// ── Main entry: AI first, pixel fallback ──
function analyzeOMRBubbles(imageData, W, H, numQuestions) {
  // This is called from processOMRImage synchronously,
  // so we return a Promise-like object that resolves async.
  // We'll monkey-patch processOMRImage to handle async below.
  return { _isAsyncOMR: true, imageData, W, H, numQuestions };
}

// ── Override processOMRImage to support async AI analysis ──
function processOMRImage(imageData, W, H) {
  const model = scannerState.selectedModel || scannerState.selectedSubject?.models[0];
  if (!model) return showToast('يرجى اختيار النموذج أولاً', 'error');

  const numQuestions = model.questions.length;
  const el = document.getElementById('scan-result-content');

  el.innerHTML = `
    <div style="text-align:center;padding:20px;">
      <div style="font-size:32px;margin-bottom:12px;">🤖</div>
      <div style="color:var(--accent);font-size:15px;font-weight:600;">جارٍ التحليل بالذكاء الاصطناعي...</div>
      <div style="color:var(--text2);font-size:12px;margin-top:6px;">يتم تحليل ورقة الإجابات عبر Claude Vision</div>
      <div style="margin-top:16px;display:flex;justify-content:center;">
        <div style="width:40px;height:4px;background:var(--accent);border-radius:2px;
                    animation:omrPulse 1s ease-in-out infinite alternate;"></div>
      </div>
    </div>
    <style>
      @keyframes omrPulse { from{opacity:.3;transform:scaleX(.5)} to{opacity:1;transform:scaleX(1)} }
    </style>`;

  // Try AI first, fallback to pixel engine
  analyzeWithAI(imageData, W, H, numQuestions)
    .catch(err => {
      console.warn('AI analysis failed, using pixel fallback:', err.message);
      el.innerHTML = `
        <div style="color:var(--warn);font-size:12px;padding:8px;margin-bottom:8px;
                    background:#ffb84f11;border-radius:6px;border:1px solid #ffb84f33;">
          ⚠️ التحليل الذكي غير متاح — جارٍ استخدام محرك الكشف التقليدي
        </div>`;
      return analyzeWithPixels(imageData, W, H, numQuestions);
    })
    .then(results => {
      scannerState.currentScannedAnswers = results;
      const pending = [];
      Object.entries(results).forEach(([q, v]) => {
        if (v === null || v === 'MULTI') pending.push(parseInt(q));
      });
      scannerState.pendingManualQuestions = pending;
      showScanResults(results, model);
      if (pending.length > 0) showManualOverride(pending, model);
      else computeAndShowScore(results, model);
    })
    .catch(err => {
      el.innerHTML = `<div style="color:var(--danger);">❌ خطأ في تحليل الصورة: ${err.message}</div>`;
      console.error('OMR fatal error:', err);
    });
}

function showScanResults(e,t){const n=t.questions.length,a=Object.values(e).filter(e=>e&&"MULTI"!==e).length,r=Object.values(e).filter(e=>null===e||"MULTI"===e).length;let s=`<div style="margin-bottom:10px;font-size:12px;color:var(--text2);">\n    تم اكتشاف <strong style="color:var(--accent3);">${a}</strong> إجابة من أصل ${n}\n    ${r>0?`| <strong style="color:var(--warn);">${r}</strong> غير واضحة`:""}\n  </div>`;s+='<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:4px;">';for(let t=1;t<=n;t++){const n=e[t];let a="var(--bg)",r="var(--text2)",o=n||"?";null===n&&(a="#ffb84f22",r="var(--warn)",o="؟"),"MULTI"===n&&(a="#ff4f6a22",r="var(--danger)",o="!!"),n&&"MULTI"!==n&&(a="#4f7cff22",r="var(--accent)"),s+=`<div style="padding:4px;background:${a};border-radius:6px;text-align:center;border:1px solid ${a};">\n      <div style="font-size:9px;color:var(--text3);">${t}</div>\n      <div style="font-size:13px;font-weight:700;color:${r};">${o}</div>\n    </div>`}s+="</div>",document.getElementById("scan-result-content").innerHTML=s}function showManualOverride(e,t){const n=document.getElementById("manual-override-panel");document.getElementById("manual-override-list").innerHTML=e.map(e=>{const n=t.questions[e-1],a=n?n.choices.map(e=>e.label):["A","B","C","D"],r="MULTI"===scannerState.currentScannedAnswers[e];return`<div style="padding:8px;background:var(--bg);border-radius:8px;margin-bottom:8px;border:1px solid var(--border);">\n      <div style="font-size:12px;color:var(--text2);margin-bottom:6px;">\n        <strong style="color:var(--warn);">س${e}:</strong>\n        ${n?n.text.slice(0,60)+(n.text.length>60?"...":""):"سؤال رقم "+e}\n        ${r?'<span style="color:var(--danger);font-size:11px;"> (إجابات متعددة)</span>':'<span style="color:var(--warn);font-size:11px;"> (لم تُكتشف)</span>'}\n      </div>\n      <div style="display:flex;gap:6px;flex-wrap:wrap;">\n        ${a.map(t=>`<button onclick="setManualAnswer(${e},'${t}',this)"\n                  style="padding:6px 14px;border-radius:6px;border:1px solid var(--border2);\n                         background:var(--bg3);color:var(--text);cursor:pointer;\n                         font-family:Cairo,sans-serif;font-weight:600;font-size:13px;\n                         transition:all .2s;"\n                  data-q="${e}" data-label="${t}">${t}</button>`).join("")}\n        <button onclick="setManualAnswer(${e},'SKIP',this)"\n                style="padding:6px 10px;border-radius:6px;border:1px solid #ffb84f44;\n                       background:#ffb84f11;color:var(--warn);cursor:pointer;\n                       font-family:Cairo,sans-serif;font-weight:600;font-size:12px;"\n                data-q="${e}" data-label="SKIP">تجاهل</button>\n      </div>\n    </div>`}).join(""),n.style.display="block"}function setManualAnswer(e,t,n){n.parentElement.querySelectorAll("button").forEach(e=>{e.style.background="var(--bg3)",e.style.color="var(--text)",e.style.borderColor="var(--border2)"}),n.style.background="SKIP"===t?"#ffb84f33":"var(--accent)",n.style.color="SKIP"===t?"var(--warn)":"#fff",n.style.borderColor="SKIP"===t?"var(--warn)":"var(--accent)",scannerState.currentScannedAnswers[e]="SKIP"===t?null:t}function confirmManualAnswers(){if(scannerState.pendingManualQuestions.filter(e=>{const t=scannerState.currentScannedAnswers[e];return null===t||"MULTI"===t}).length>0){const e=scannerState.pendingManualQuestions.filter(e=>"MULTI"===scannerState.currentScannedAnswers[e]);if(e.length>0)return void showToast(`يرجى تحديد إجابة للأسئلة: ${e.join(", ")}`,"error")}const e=scannerState.selectedModel||scannerState.selectedSubject?.models[scannerState.selectedModelIdx||0];e&&(document.getElementById("manual-override-panel").style.display="none",computeAndShowScore(scannerState.currentScannedAnswers,e))}function computeAndShowScore(e,t){let n=0;const a=t.questions.length;t.questions.forEach((t,a)=>{const r=e[a+1];r&&r===t.correctLabel&&n++});const r=document.getElementById("score-summary-panel");document.getElementById("score-display").textContent=`${n} / ${a}`,document.getElementById("score-percent").textContent=`${Math.round(n/a*100)}%`,r.style.display="block",scannerState._pendingResult={studentId:scannerState.currentStudentId,modelName:t.name,modelLetter:t.letter,subject:scannerState.selectedSubject?.name||"",answers:{...e},correct:n,total:a,score:n,percent:Math.round(n/a*100),timestamp:(new Date).toISOString()}}function saveAndNextStudent(){scannerState._pendingResult&&(scannerState.allResults=scannerState.allResults.filter(e=>!(e.studentId===scannerState._pendingResult.studentId&&e.subject===scannerState._pendingResult.subject)),scannerState.allResults.push(scannerState._pendingResult),saveScanResults(scannerState.allResults),renderResultsTable(),showToast(`تم حفظ نتيجة الطالب ${scannerState._pendingResult.studentId} ✓`,"success"),scannerState._pendingResult=null,scannerState.currentStudentId=null,scannerState.currentScannedAnswers=null,scannerState.pendingManualQuestions=[],document.getElementById("scanner-current-student").textContent="-",document.getElementById("scanner-student-id").value="",document.getElementById("scanner-camera-card").style.display="none",document.getElementById("scan-result-content").textContent="اضغط على زر المسح لبدء العملية...",document.getElementById("manual-override-panel").style.display="none",document.getElementById("score-summary-panel").style.display="none",stopCamera(),setTimeout(()=>document.getElementById("scanner-student-id").focus(),100))}function rescanCurrent(){scannerState.currentScannedAnswers=null,scannerState.pendingManualQuestions=[],scannerState._pendingResult=null,document.getElementById("scan-result-content").textContent="اضغط على زر المسح لبدء العملية...",document.getElementById("manual-override-panel").style.display="none",document.getElementById("score-summary-panel").style.display="none",scannerState.cameraStream||startCamera()}function renderResultsTable(){const e=document.getElementById("scanner-results-card"),t=document.getElementById("results-tbody"),n=document.getElementById("results-stats");if(!t)return;const a=scannerState.selectedSubject?.name||"",r=a?scannerState.allResults.filter(e=>e.subject===a):scannerState.allResults;if(0===r.length)return void(e.style.display="none");e.style.display="block",t.innerHTML=r.map((e,t)=>{const n=e.percent??Math.round(e.correct/e.total*100),a=n>=60?"var(--accent3)":n>=50?"var(--warn)":"var(--danger)";return`<tr>\n      <td>${t+1}</td>\n      <td><strong style="color:var(--accent);">${e.studentId}</strong></td>\n      <td>${e.modelName||e.modelLetter||"-"}</td>\n      <td>${e.correct} / ${e.total}</td>\n      <td><strong style="color:${a};">${e.score}</strong></td>\n      <td><span style="color:${a};">${n}%</span></td>\n      <td>\n        <button onclick="deleteResult('${e.studentId}','${e.subject}')"\n                style="padding:4px 10px;background:#ff4f6a22;color:var(--danger);border:none;\n                       border-radius:4px;cursor:pointer;font-family:Cairo,sans-serif;font-size:11px;">حذف</button>\n      </td>\n    </tr>`}).join("");const s=r.length>0?Math.round(r.reduce((e,t)=>e+(t.percent??0),0)/r.length):0,o=r.filter(e=>(e.percent??0)>=60).length;n.innerHTML=`إجمالي الطلاب: <strong>${r.length}</strong> | متوسط الدرجات: <strong style="color:var(--accent);">${s}%</strong> | الناجحون (≥60%): <strong style="color:var(--accent3);">${o}</strong>`}function deleteResult(e,t){confirm(`حذف نتيجة الطالب ${e}؟`)&&(scannerState.allResults=scannerState.allResults.filter(n=>!(n.studentId===e&&n.subject===t)),saveScanResults(scannerState.allResults),renderResultsTable(),showToast("تم حذف النتيجة","success"))}function clearAllResults(){if(!confirm("هل أنت متأكد من حذف جميع النتائج؟ لا يمكن التراجع عن هذا الإجراء."))return;const e=scannerState.selectedSubject?.name||"";scannerState.allResults=e?scannerState.allResults.filter(t=>t.subject!==e):[],saveScanResults(scannerState.allResults),renderResultsTable(),showToast("تم مسح النتائج","success")}function exportResultsExcel(){const e=scannerState.selectedSubject?.name||"",t=e?scannerState.allResults.filter(t=>t.subject===e):scannerState.allResults;if(0===t.length)return void showToast("لا توجد نتائج للتصدير","error");const n=XLSX.utils.book_new(),a=[["الرقم الجامعي","النموذج","المادة","الإجابات الصحيحة","مجموع الأسئلة","العلامة","النسبة المئوية","تاريخ المسح"]];t.forEach(e=>{const t=e.percent??Math.round(e.correct/e.total*100);a.push([e.studentId,e.modelName||e.modelLetter||"",e.subject,e.correct,e.total,e.score,t+"%",e.timestamp?new Date(e.timestamp).toLocaleString("ar-EG"):""])});const r=XLSX.utils.aoa_to_sheet(a);r["!cols"]=[{wch:18},{wch:14},{wch:20},{wch:16},{wch:14},{wch:10},{wch:14},{wch:22}],XLSX.utils.book_append_sheet(n,r,"نتائج الامتحان");const s=`نتائج_${e||"الامتحان"}_${(new Date).toLocaleDateString("ar-EG").replace(/\//g,"-")}.xlsx`;XLSX.writeFile(n,s),showToast("تم تحميل ملف Excel ✓","success")}
