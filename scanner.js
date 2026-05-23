const RESULTS_STORAGE_KEY="omr_results_v1";let scannerState={selectedSubject:null,selectedModelIdx:null,selectedModel:null,currentStudentId:null,currentScannedAnswers:null,pendingManualQuestions:[],cameraStream:null,facingMode:"environment",allResults:[]};function loadScanResults(){try{const e=localStorage.getItem("omr_results_v1");return e?JSON.parse(e):[]}catch(e){return[]}}function saveScanResults(e){try{localStorage.setItem("omr_results_v1",JSON.stringify(e))}catch(e){console.error("Save results error:",e)}}function initScannerPanel(){const e=getAllProjects(),t=document.getElementById("scanner-subject-select");if(t){if(t.innerHTML='<option value="">-- اختر المادة --</option>',models.length>0&&examCfg.subject){const e=document.createElement("option");e.value="__current__",e.textContent=`${examCfg.subject} (الجلسة الحالية — ${models.length} نماذج)`,t.appendChild(e)}e.forEach(e=>{const n=document.createElement("option");n.value=e.id,n.textContent=`${e.name} (${e.models} نماذج — ${new Date(e.date).toLocaleDateString("ar-EG")})`,t.appendChild(n)}),scannerState.allResults=loadScanResults(),scannerState.allResults.length>0&&renderResultsTable()}}function onScannerSubjectChange(){const e=document.getElementById("scanner-subject-select").value,t=document.getElementById("scanner-model-select"),n=document.getElementById("scanner-student-card"),a=document.getElementById("scanner-model-info");if(t.innerHTML='<option value="">-- اختر النموذج --</option>',n.style.display="none",a.style.display="none",scannerState.selectedSubject=null,!e)return;let r=[],s="";if("__current__"===e)r=models,s=examCfg.subject;else{const t=loadProjectFromStorage(parseInt(e));t&&(r=t.models||[],s=t.examCfg?.subject||"غير محدد")}if(0===r.length)return void showToast("لا توجد نماذج لهذه المادة","error");scannerState.selectedSubject={val:e,models:r,name:s},r.forEach((e,n)=>{const a=document.createElement("option");a.value=n,a.textContent=`${e.name} — ${e.questions.length} سؤال`,t.appendChild(a)});const o=document.createElement("option");o.value="auto",o.textContent="🔍 تحديد تلقائي (من الورقة)",t.insertBefore(o,t.children[1]),n.style.display="block",document.getElementById("results-subject-label").textContent=s}function onScannerModelChange(){const e=document.getElementById("scanner-model-select").value,t=document.getElementById("scanner-model-info");if(!e||!scannerState.selectedSubject)return void(t.style.display="none");if("auto"===e)return t.style.display="block",t.innerHTML="🔍 سيتم تحديد النموذج تلقائياً من الورقة. إذا لم يتمكن الماسح من تحديده، سيطلب منك الاختيار يدوياً.",void(scannerState.selectedModelIdx="auto");const n=parseInt(e),a=scannerState.selectedSubject.models[n];scannerState.selectedModelIdx=n,scannerState.selectedModel=a,t.style.display="block",t.innerHTML=`✅ النموذج: <strong style="color:var(--accent);">${a.name}</strong> | عدد الأسئلة: <strong>${a.questions.length}</strong> | مفتاح الإجابات جاهز للمقارنة`}function startScanSession(){const e=document.getElementById("scanner-student-id").value.trim();if(e)if(document.getElementById("scanner-model-select").value){if(scannerState.allResults.find(t=>t.studentId===e&&t.subject===scannerState.selectedSubject?.name)){if(!confirm(`الرقم الجامعي "${e}" موجود مسبقاً لهذه المادة. هل تريد الاستبدال؟`))return;scannerState.allResults=scannerState.allResults.filter(t=>!(t.studentId===e&&t.subject===scannerState.selectedSubject?.name))}scannerState.currentStudentId=e,document.getElementById("scanner-current-student").textContent=e,document.getElementById("scanner-camera-card").style.display="block",document.getElementById("scan-result-content").textContent="اضغط على زر المسح لبدء العملية...",document.getElementById("manual-override-panel").style.display="none",document.getElementById("score-summary-panel").style.display="none",startCamera()}else showToast("يرجى اختيار النموذج","error");else showToast("يرجى إدخال الرقم الجامعي","error")}async function startCamera(){try{scannerState.cameraStream&&scannerState.cameraStream.getTracks().forEach(e=>e.stop());const e=await navigator.mediaDevices.getUserMedia({video:{facingMode:scannerState.facingMode,width:{ideal:1280},height:{ideal:960}}});scannerState.cameraStream=e,document.getElementById("scanner-video").srcObject=e}catch(e){showToast("لا يمكن الوصول للكاميرا: "+e.message,"error")}}function stopCamera(){scannerState.cameraStream&&(scannerState.cameraStream.getTracks().forEach(e=>e.stop()),scannerState.cameraStream=null),document.getElementById("scanner-video").srcObject=null}function switchCamera(){scannerState.facingMode="environment"===scannerState.facingMode?"user":"environment",startCamera()}function captureAndScan(){const e=document.getElementById("scanner-video");if(!e.srcObject)return void showToast("الكاميرا غير مفعّلة","error");const t=document.createElement("canvas");t.width=e.videoWidth||640,t.height=e.videoHeight||480;const n=t.getContext("2d");n.drawImage(e,0,0,t.width,t.height),processOMRImage(n.getImageData(0,0,t.width,t.height),t.width,t.height)}function scanFromFile(e){const t=e.target.files[0];if(!t)return;const n=new Image;n.onload=()=>{const e=document.createElement("canvas");e.width=n.width,e.height=n.height;const t=e.getContext("2d");t.drawImage(n,0,0),processOMRImage(t.getImageData(0,0,e.width,e.height),e.width,e.height)},n.src=URL.createObjectURL(t)}function processOMRImage(e,t,n){const a=scannerState.selectedModel||scannerState.selectedSubject?.models[0];a?(document.getElementById("scan-result-content").innerHTML='<div style="color:var(--accent);">⏳ جارٍ تحليل الصورة...</div>',setTimeout(()=>{try{const r=analyzeOMRBubbles(e,t,n,a.questions.length);scannerState.currentScannedAnswers=r;const s=[];Object.entries(r).forEach(([e,t])=>{null!==t&&"MULTI"!==t||s.push(parseInt(e))}),scannerState.pendingManualQuestions=s,showScanResults(r,a),s.length>0?showManualOverride(s,a):computeAndShowScore(r,a)}catch(e){document.getElementById("scan-result-content").innerHTML=`<div style="color:var(--danger);">❌ خطأ في تحليل الصورة: ${e.message}</div>`,console.error("OMR error:",e)}},100)):showToast("يرجى اختيار النموذج أولاً","error")}function analyzeOMRBubbles(imageData, W, H, numQuestions) {
  const data = imageData.data;
  const gray = new Uint8Array(W * H);

  // Step 1: Grayscale
  for (let i = 0; i < W * H; i++) {
    gray[i] = Math.round(0.299 * data[4*i] + 0.587 * data[4*i+1] + 0.114 * data[4*i+2]);
  }

  // Step 2: Adaptive threshold - find Otsu threshold
  const hist = new Array(256).fill(0);
  for (let i = 0; i < gray.length; i++) hist[gray[i]]++;
  let total = gray.length, sumB = 0, wB = 0, sum = 0, max = 0, threshold = 128;
  for (let i = 0; i < 256; i++) sum += i * hist[i];
  for (let t = 0; t < 256; t++) {
    wB += hist[t]; if (!wB) continue;
    const wF = total - wB; if (!wF) break;
    sumB += t * hist[t];
    const mB = sumB / wB, mF = (sum - sumB) / wF;
    const between = wB * wF * (mB - mF) ** 2;
    if (between > max) { max = between; threshold = t; }
  }
  // Binary: true = dark (filled bubble)
  const binary = new Uint8Array(W * H);
  for (let i = 0; i < gray.length; i++) binary[i] = gray[i] < threshold ? 1 : 0;

  // Step 3: Row/Col density profiles to find table boundaries
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

  // Find active region (where density > threshold)
  const rowTh = 0.05, colTh = 0.03;
  let rTop = Math.floor(H * 0.05), rBot = Math.floor(H * 0.95);
  let cLeft = Math.floor(W * 0.02), cRight = Math.floor(W * 0.98);

  for (let y = Math.floor(H * 0.05); y < H * 0.95; y++) if (rowDark[y] >= rowTh) { rTop = y; break; }
  for (let y = Math.floor(H * 0.95); y > rTop; y--) if (rowDark[y] >= rowTh) { rBot = y; break; }
  for (let x = Math.floor(W * 0.02); x < W * 0.98; x++) if (colDark[x] >= colTh) { cLeft = x; break; }
  for (let x = Math.floor(W * 0.98); x > cLeft; x--) if (colDark[x] >= colTh) { cRight = x; break; }

  const tableH = rBot - rTop;
  const tableW = cRight - cLeft;

  // Step 4: Detect grid lines using column profile peaks
  // Find 5 answer columns (A,B,C,D + question label col) by looking at vertical line peaks
  // Instead: divide evenly but skip first column (question label)
  // Assume layout: col0=label, col1=A, col2=B, col3=C, col4=D
  // and rows: row0=header, row1..N = questions

  // Detect horizontal grid lines (row separators) using row dark density peaks
  // Grid lines appear as rows with high density
  const gridLineTh = 0.3;
  const gridRows = [];
  let inLine = false;
  for (let y = rTop; y <= rBot; y++) {
    if (rowDark[y] >= gridLineTh) {
      if (!inLine) { gridRows.push(y); inLine = true; }
    } else { inLine = false; }
  }

  // Detect vertical grid lines
  const gridLineTh2 = 0.2;
  const gridCols = [];
  inLine = false;
  for (let x = cLeft; x <= cRight; x++) {
    if (colDark[x] >= gridLineTh2) {
      if (!inLine) { gridCols.push(x); inLine = true; }
    } else { inLine = false; }
  }

  // Build cell regions from grid lines
  // Filter: need at least (numQuestions+1) row lines and at least 5 col lines
  // Fallback to even division if grid detection fails
  let rowBounds = [], colBounds = [];

  if (gridRows.length >= 3) {
    // merge nearby lines (within 5px)
    const mergedRows = [gridRows[0]];
    for (let i = 1; i < gridRows.length; i++) {
      if (gridRows[i] - mergedRows[mergedRows.length-1] > 8) mergedRows.push(gridRows[i]);
    }
    rowBounds = mergedRows;
  }

  if (gridCols.length >= 4) {
    const mergedCols = [gridCols[0]];
    for (let i = 1; i < gridCols.length; i++) {
      if (gridCols[i] - mergedCols[mergedCols.length-1] > 8) mergedCols.push(gridCols[i]);
    }
    colBounds = mergedCols;
  }

  // Fallback: even division
  const cellH = tableH / (numQuestions + 2); // +2 for header rows
  const cellW = tableW / 6; // 1 label + 5 options (but usually 4 choices)

  // Build question row ranges
  const qRows = [];
  if (rowBounds.length >= 3) {
    // Use detected grid lines as row separators
    // Skip first 1-2 rows (header)
    const headerRows = rowBounds.length > numQuestions + 1 ? rowBounds.length - numQuestions - 1 : 1;
    for (let i = headerRows; i < rowBounds.length && qRows.length < numQuestions; i++) {
      const y0 = rowBounds[i];
      const y1 = i + 1 < rowBounds.length ? rowBounds[i+1] : rBot;
      if (y1 - y0 > 10) qRows.push([y0, y1]);
    }
  }
  // Fallback
  if (qRows.length < numQuestions) {
    qRows.length = 0;
    const skip = rowBounds.length >= 2 ? rowBounds[1] - rTop : cellH * 2;
    const startY = rTop + skip;
    const usableH = rBot - startY;
    const qCellH = usableH / numQuestions;
    for (let q = 0; q < numQuestions; q++) {
      qRows.push([Math.floor(startY + q * qCellH), Math.floor(startY + (q+1) * qCellH)]);
    }
  }

  // Build answer column ranges (skip first col = label)
  const numChoices = 4;
  const aCols = [];
  if (colBounds.length >= 5) {
    // Skip first column boundary (label col)
    const labelEnd = colBounds.length > numChoices ? colBounds[colBounds.length - numChoices - 1] : colBounds[0];
    const answerCols = colBounds.filter(x => x > labelEnd);
    for (let c = 0; c < Math.min(numChoices, answerCols.length); c++) {
      const x0 = answerCols[c];
      const x1 = c + 1 < answerCols.length ? answerCols[c+1] : cRight;
      aCols.push([x0, x1]);
    }
  }
  // Fallback: even division, skip ~30% for label col
  if (aCols.length < numChoices) {
    aCols.length = 0;
    const labelColW = tableW * 0.30;
    const ansStart = cLeft + labelColW;
    const ansW = (cRight - ansStart) / numChoices;
    for (let c = 0; c < numChoices; c++) {
      aCols.push([Math.floor(ansStart + c * ansW), Math.floor(ansStart + (c+1) * ansW)]);
    }
  }

  // Step 5: For each question row x each choice col, compute fill ratio
  const choiceLabels = [A, B, C, D];
  const results = {};
  const MARGIN = 0.15; // fraction of cell to ignore (border pixels)

  for (let q = 0; q < numQuestions; q++) {
    const [ry0, ry1] = qRows[q] || [0, 0];
    const rowH2 = ry1 - ry0;
    const ratios = [];

    for (let c = 0; c < numChoices; c++) {
      if (c >= aCols.length) { ratios.push(0); continue; }
      const [cx0, cx1] = aCols[c];
      const colW2 = cx1 - cx0;
      const mx = Math.max(2, Math.floor(colW2 * MARGIN));
      const my = Math.max(2, Math.floor(rowH2 * MARGIN));

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

    // Normalize: find max ratio
    const maxRatio = Math.max(...ratios);

    // Dynamic threshold: filled bubble should be significantly darker than empty
    // Use relative threshold: filled = maxRatio, consider filled if > 50% of max
    // AND absolute threshold: must have at least some darkness
    const absTh = 0.08;  // absolute minimum to be filled
    const relTh = 0.5;   // must be >= 50% of the darkest bubble in this row

    const candidates = ratios
      .map((r, i) => ({ label: choiceLabels[i], ratio: r }))
      .filter(x => x.ratio >= absTh && x.ratio >= maxRatio * relTh);

    if (candidates.length === 0) {
      results[q + 1] = null;
    } else if (candidates.length === 1) {
      results[q + 1] = candidates[0].label;
    } else {
      candidates.sort((a, b) => b.ratio - a.ratio);
      const gap = candidates[0].ratio - candidates[1].ratio;
      // If top two are very close in fill, mark as MULTI
      const closeThreshold = candidates[0].ratio * 0.20;
      results[q + 1] = gap < closeThreshold ? MULTI : candidates[0].label;
    }
  }

  return results;
}
function showScanResults(e,t){const n=t.questions.length,a=Object.values(e).filter(e=>e&&"MULTI"!==e).length,r=Object.values(e).filter(e=>null===e||"MULTI"===e).length;let s=`<div style="margin-bottom:10px;font-size:12px;color:var(--text2);">\n    تم اكتشاف <strong style="color:var(--accent3);">${a}</strong> إجابة من أصل ${n}\n    ${r>0?`| <strong style="color:var(--warn);">${r}</strong> غير واضحة`:""}\n  </div>`;s+='<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:4px;">';for(let t=1;t<=n;t++){const n=e[t];let a="var(--bg)",r="var(--text2)",o=n||"?";null===n&&(a="#ffb84f22",r="var(--warn)",o="؟"),"MULTI"===n&&(a="#ff4f6a22",r="var(--danger)",o="!!"),n&&"MULTI"!==n&&(a="#4f7cff22",r="var(--accent)"),s+=`<div style="padding:4px;background:${a};border-radius:6px;text-align:center;border:1px solid ${a};">\n      <div style="font-size:9px;color:var(--text3);">${t}</div>\n      <div style="font-size:13px;font-weight:700;color:${r};">${o}</div>\n    </div>`}s+="</div>",document.getElementById("scan-result-content").innerHTML=s}function showManualOverride(e,t){const n=document.getElementById("manual-override-panel");document.getElementById("manual-override-list").innerHTML=e.map(e=>{const n=t.questions[e-1],a=n?n.choices.map(e=>e.label):["A","B","C","D"],r="MULTI"===scannerState.currentScannedAnswers[e];return`<div style="padding:8px;background:var(--bg);border-radius:8px;margin-bottom:8px;border:1px solid var(--border);">\n      <div style="font-size:12px;color:var(--text2);margin-bottom:6px;">\n        <strong style="color:var(--warn);">س${e}:</strong>\n        ${n?n.text.slice(0,60)+(n.text.length>60?"...":""):"سؤال رقم "+e}\n        ${r?'<span style="color:var(--danger);font-size:11px;"> (إجابات متعددة)</span>':'<span style="color:var(--warn);font-size:11px;"> (لم تُكتشف)</span>'}\n      </div>\n      <div style="display:flex;gap:6px;flex-wrap:wrap;">\n        ${a.map(t=>`<button onclick="setManualAnswer(${e},'${t}',this)"\n                  style="padding:6px 14px;border-radius:6px;border:1px solid var(--border2);\n                         background:var(--bg3);color:var(--text);cursor:pointer;\n                         font-family:Cairo,sans-serif;font-weight:600;font-size:13px;\n                         transition:all .2s;"\n                  data-q="${e}" data-label="${t}">${t}</button>`).join("")}\n        <button onclick="setManualAnswer(${e},'SKIP',this)"\n                style="padding:6px 10px;border-radius:6px;border:1px solid #ffb84f44;\n                       background:#ffb84f11;color:var(--warn);cursor:pointer;\n                       font-family:Cairo,sans-serif;font-weight:600;font-size:12px;"\n                data-q="${e}" data-label="SKIP">تجاهل</button>\n      </div>\n    </div>`}).join(""),n.style.display="block"}function setManualAnswer(e,t,n){n.parentElement.querySelectorAll("button").forEach(e=>{e.style.background="var(--bg3)",e.style.color="var(--text)",e.style.borderColor="var(--border2)"}),n.style.background="SKIP"===t?"#ffb84f33":"var(--accent)",n.style.color="SKIP"===t?"var(--warn)":"#fff",n.style.borderColor="SKIP"===t?"var(--warn)":"var(--accent)",scannerState.currentScannedAnswers[e]="SKIP"===t?null:t}function confirmManualAnswers(){if(scannerState.pendingManualQuestions.filter(e=>{const t=scannerState.currentScannedAnswers[e];return null===t||"MULTI"===t}).length>0){const e=scannerState.pendingManualQuestions.filter(e=>"MULTI"===scannerState.currentScannedAnswers[e]);if(e.length>0)return void showToast(`يرجى تحديد إجابة للأسئلة: ${e.join(", ")}`,"error")}const e=scannerState.selectedModel||scannerState.selectedSubject?.models[scannerState.selectedModelIdx||0];e&&(document.getElementById("manual-override-panel").style.display="none",computeAndShowScore(scannerState.currentScannedAnswers,e))}function computeAndShowScore(e,t){let n=0;const a=t.questions.length;t.questions.forEach((t,a)=>{const r=e[a+1];r&&r===t.correctLabel&&n++});const r=document.getElementById("score-summary-panel");document.getElementById("score-display").textContent=`${n} / ${a}`,document.getElementById("score-percent").textContent=`${Math.round(n/a*100)}%`,r.style.display="block",scannerState._pendingResult={studentId:scannerState.currentStudentId,modelName:t.name,modelLetter:t.letter,subject:scannerState.selectedSubject?.name||"",answers:{...e},correct:n,total:a,score:n,percent:Math.round(n/a*100),timestamp:(new Date).toISOString()}}function saveAndNextStudent(){scannerState._pendingResult&&(scannerState.allResults=scannerState.allResults.filter(e=>!(e.studentId===scannerState._pendingResult.studentId&&e.subject===scannerState._pendingResult.subject)),scannerState.allResults.push(scannerState._pendingResult),saveScanResults(scannerState.allResults),renderResultsTable(),showToast(`تم حفظ نتيجة الطالب ${scannerState._pendingResult.studentId} ✓`,"success"),scannerState._pendingResult=null,scannerState.currentStudentId=null,scannerState.currentScannedAnswers=null,scannerState.pendingManualQuestions=[],document.getElementById("scanner-current-student").textContent="-",document.getElementById("scanner-student-id").value="",document.getElementById("scanner-camera-card").style.display="none",document.getElementById("scan-result-content").textContent="اضغط على زر المسح لبدء العملية...",document.getElementById("manual-override-panel").style.display="none",document.getElementById("score-summary-panel").style.display="none",stopCamera(),setTimeout(()=>document.getElementById("scanner-student-id").focus(),100))}function rescanCurrent(){scannerState.currentScannedAnswers=null,scannerState.pendingManualQuestions=[],scannerState._pendingResult=null,document.getElementById("scan-result-content").textContent="اضغط على زر المسح لبدء العملية...",document.getElementById("manual-override-panel").style.display="none",document.getElementById("score-summary-panel").style.display="none",scannerState.cameraStream||startCamera()}function renderResultsTable(){const e=document.getElementById("scanner-results-card"),t=document.getElementById("results-tbody"),n=document.getElementById("results-stats");if(!t)return;const a=scannerState.selectedSubject?.name||"",r=a?scannerState.allResults.filter(e=>e.subject===a):scannerState.allResults;if(0===r.length)return void(e.style.display="none");e.style.display="block",t.innerHTML=r.map((e,t)=>{const n=e.percent??Math.round(e.correct/e.total*100),a=n>=60?"var(--accent3)":n>=50?"var(--warn)":"var(--danger)";return`<tr>\n      <td>${t+1}</td>\n      <td><strong style="color:var(--accent);">${e.studentId}</strong></td>\n      <td>${e.modelName||e.modelLetter||"-"}</td>\n      <td>${e.correct} / ${e.total}</td>\n      <td><strong style="color:${a};">${e.score}</strong></td>\n      <td><span style="color:${a};">${n}%</span></td>\n      <td>\n        <button onclick="deleteResult('${e.studentId}','${e.subject}')"\n                style="padding:4px 10px;background:#ff4f6a22;color:var(--danger);border:none;\n                       border-radius:4px;cursor:pointer;font-family:Cairo,sans-serif;font-size:11px;">حذف</button>\n      </td>\n    </tr>`}).join("");const s=r.length>0?Math.round(r.reduce((e,t)=>e+(t.percent??0),0)/r.length):0,o=r.filter(e=>(e.percent??0)>=60).length;n.innerHTML=`إجمالي الطلاب: <strong>${r.length}</strong> | متوسط الدرجات: <strong style="color:var(--accent);">${s}%</strong> | الناجحون (≥60%): <strong style="color:var(--accent3);">${o}</strong>`}function deleteResult(e,t){confirm(`حذف نتيجة الطالب ${e}؟`)&&(scannerState.allResults=scannerState.allResults.filter(n=>!(n.studentId===e&&n.subject===t)),saveScanResults(scannerState.allResults),renderResultsTable(),showToast("تم حذف النتيجة","success"))}function clearAllResults(){if(!confirm("هل أنت متأكد من حذف جميع النتائج؟ لا يمكن التراجع عن هذا الإجراء."))return;const e=scannerState.selectedSubject?.name||"";scannerState.allResults=e?scannerState.allResults.filter(t=>t.subject!==e):[],saveScanResults(scannerState.allResults),renderResultsTable(),showToast("تم مسح النتائج","success")}function exportResultsExcel(){const e=scannerState.selectedSubject?.name||"",t=e?scannerState.allResults.filter(t=>t.subject===e):scannerState.allResults;if(0===t.length)return void showToast("لا توجد نتائج للتصدير","error");const n=XLSX.utils.book_new(),a=[["الرقم الجامعي","النموذج","المادة","الإجابات الصحيحة","مجموع الأسئلة","العلامة","النسبة المئوية","تاريخ المسح"]];t.forEach(e=>{const t=e.percent??Math.round(e.correct/e.total*100);a.push([e.studentId,e.modelName||e.modelLetter||"",e.subject,e.correct,e.total,e.score,t+"%",e.timestamp?new Date(e.timestamp).toLocaleString("ar-EG"):""])});const r=XLSX.utils.aoa_to_sheet(a);r["!cols"]=[{wch:18},{wch:14},{wch:20},{wch:16},{wch:14},{wch:10},{wch:14},{wch:22}],XLSX.utils.book_append_sheet(n,r,"نتائج الامتحان");const s=`نتائج_${e||"الامتحان"}_${(new Date).toLocaleDateString("ar-EG").replace(/\//g,"-")}.xlsx`;XLSX.writeFile(n,s),showToast("تم تحميل ملف Excel ✓","success")}
