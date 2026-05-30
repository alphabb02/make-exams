const SCANNER_RESULTS_STORAGE_KEY="omr_results_v1";let scannerModuleState={selectedSubject:null,selectedModelIdx:null,selectedModel:null,currentStudentId:null,currentScannedAnswers:null,pendingManualQuestions:[],cameraStream:null,facingMode:"environment",allResults:[]};function loadScanResults(){try{const e=localStorage.getItem(SCANNER_RESULTS_STORAGE_KEY);return e?JSON.parse(e):[]}catch(e){return[]}}function saveScanResults(e){try{localStorage.setItem(SCANNER_RESULTS_STORAGE_KEY,JSON.stringify(e))}catch(e){console.error("Save results error:",e)}}function initScannerPanel(){const e=getAllProjects(),t=document.getElementById("scanner-subject-select");if(t){if(t.innerHTML='<option value="">-- اختر المادة --</option>',models.length>0&&examCfg.subject){const e=document.createElement("option");e.value="__current__",e.textContent=`${examCfg.subject} (الجلسة الحالية — ${models.length} نماذج)`,t.appendChild(e)}e.forEach(e=>{const n=document.createElement("option");n.value=e.id,n.textContent=`${e.name} (${e.models} نماذج — ${new Date(e.date).toLocaleDateString("ar-EG")})`,t.appendChild(n)}),scannerModuleState.allResults=loadScanResults(),scannerModuleState.allResults.length>0&&renderResultsTable()}}function onScannerSubjectChange(){const e=document.getElementById("scanner-subject-select").value,t=document.getElementById("scanner-model-select"),n=document.getElementById("scanner-student-card"),a=document.getElementById("scanner-model-info");if(t.innerHTML='<option value="">-- اختر النموذج --</option>',n.style.display="none",a.style.display="none",scannerModuleState.selectedSubject=null,!e)return;let r=[],s="";if("__current__"===e)r=models,s=examCfg.subject;else{const t=loadProjectFromStorage(parseInt(e));t&&(r=t.models||[],s=t.examCfg?.subject||"غير محدد")}if(0===r.length)return void showToast("لا توجد نماذج لهذه المادة","error");scannerModuleState.selectedSubject={val:e,models:r,name:s},r.forEach((e,n)=>{const a=document.createElement("option");a.value=n,a.textContent=`${e.name} — ${e.questions.length} سؤال`,t.appendChild(a)});const o=document.createElement("option");o.value="auto",o.textContent="🔍 تحديد تلقائي (من الورقة)",t.insertBefore(o,t.children[1]),n.style.display="block",document.getElementById("results-subject-label").textContent=s}function onScannerModelChange(){const e=document.getElementById("scanner-model-select").value,t=document.getElementById("scanner-model-info");if(!e||!scannerModuleState.selectedSubject)return void(t.style.display="none");if("auto"===e)return t.style.display="block",t.innerHTML="🔍 سيتم تحديد النموذج تلقائياً من الورقة. إذا لم يتمكن الماسح من تحديده، سيطلب منك الاختيار يدوياً.",void(scannerModuleState.selectedModelIdx="auto");const n=parseInt(e),a=scannerModuleState.selectedSubject.models[n];scannerModuleState.selectedModelIdx=n,scannerModuleState.selectedModel=a,t.style.display="block",t.innerHTML=`✅ النموذج: <strong style="color:var(--accent);">${a.name}</strong> | عدد الأسئلة: <strong>${a.questions.length}</strong> | مفتاح الإجابات جاهز للمقارنة`}function startScanSession(){const e=document.getElementById("scanner-student-id").value.trim();if(e)if(document.getElementById("scanner-model-select").value){if(scannerModuleState.allResults.find(t=>t.studentId===e&&t.subject===scannerModuleState.selectedSubject?.name)){if(!confirm(`الرقم الجامعي "${e}" موجود مسبقاً لهذه المادة. هل تريد الاستبدال؟`))return;scannerModuleState.allResults=scannerModuleState.allResults.filter(t=>!(t.studentId===e&&t.subject===scannerModuleState.selectedSubject?.name))}scannerModuleState.currentStudentId=e,document.getElementById("scanner-current-student").textContent=e,document.getElementById("scanner-camera-card").style.display="block",document.getElementById("scan-result-content").textContent="اضغط على زر المسح لبدء العملية...",document.getElementById("manual-override-panel").style.display="none",document.getElementById("score-summary-panel").style.display="none",startCamera()}else showToast("يرجى اختيار النموذج","error");else showToast("يرجى إدخال الرقم الجامعي","error")}async function startCamera(){try{scannerModuleState.cameraStream&&scannerModuleState.cameraStream.getTracks().forEach(e=>e.stop());const e=await navigator.mediaDevices.getUserMedia({video:{facingMode:scannerModuleState.facingMode,width:{ideal:1280},height:{ideal:960}}});scannerModuleState.cameraStream=e,document.getElementById("scanner-video").srcObject=e}catch(e){showToast("لا يمكن الوصول للكاميرا: "+e.message,"error")}}function stopCamera(){scannerModuleState.cameraStream&&(scannerModuleState.cameraStream.getTracks().forEach(e=>e.stop()),scannerModuleState.cameraStream=null),document.getElementById("scanner-video").srcObject=null}function switchCamera(){scannerModuleState.facingMode="environment"===scannerModuleState.facingMode?"user":"environment",startCamera()}function captureAndScan(){const e=document.getElementById("scanner-video");if(!e.srcObject)return void showToast("الكاميرا غير مفعّلة","error");const t=document.createElement("canvas");t.width=e.videoWidth||640,t.height=e.videoHeight||480;const n=t.getContext("2d");n.drawImage(e,0,0,t.width,t.height),processOMRImage(n.getImageData(0,0,t.width,t.height),t.width,t.height)}function scanFromFile(e){const t=e.target.files[0];if(!t)return;const n=new Image;n.onload=()=>{const e=document.createElement("canvas");e.width=n.width,e.height=n.height;const t=e.getContext("2d");t.drawImage(n,0,0),processOMRImage(t.getImageData(0,0,e.width,e.height),e.width,e.height)},n.src=URL.createObjectURL(t)}// ============================================================
//  AI-POWERED OMR ENGINE  —  analyzeOMRBubbles + processOMRImage
//  Uses Google Gemini Flash as primary engine,
//  with a robust pixel-based fallback.
// ============================================================

// ── Helper: convert ImageData → base64 JPEG via offscreen canvas ──
function imageDataToBase64(imageData, W, H) {
  const maxDim = 640;
  const scale = Math.min(1, maxDim / Math.max(W, H));
  const width = Math.max(1, Math.round(W * scale));
  const height = Math.max(1, Math.round(H * scale));

  const srcCanvas = document.createElement('canvas');
  srcCanvas.width = W;
  srcCanvas.height = H;
  srcCanvas.getContext('2d').putImageData(imageData, 0, 0);

  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');
  ctx.drawImage(srcCanvas, 0, 0, width, height);

  return canvas.toDataURL('image/jpeg', 0.7).split(',')[1];
}

// ── Build the AI prompt ──
function buildOMRPrompt(numQuestions) {
  return `
You are an expert OMR scanner. Analyze the provided image of an answer sheet.
The answer grid contains 4 rows of bubbles from top to bottom: A, B, C, D.
Columns represent questions from left to right, from 1 to ${numQuestions}.
Task: Extract selected options for each question.
Rules:
1. Return ONLY a valid JSON object. Do not include markdown (\`\`\`json), explanations, or extra text.
2. Format exactly: {"1": ["A"], "2": ["B", "C"], "3": null, ...}.
3. If a bubble is clearly filled, include it. If multiple bubbles are clearly filled, list them all.
4. If you are not confident about a question, return null for that question.
5. Use only values "A", "B", "C", "D", or null.
6. Ignore headers, labels, or any text outside the answer grid.
7. Do not add any properties outside the question numbers 1..${numQuestions}.
`;
}

// Keep the API key on the server-side (proxy). Do NOT store it in the client.
const GOOGLE_GEMINI_MODEL = 'gemini-flash-latest';
// Use the remote worker proxy provided by the user. This forwards requests to
// Google's Generative API so the client doesn't need the API key.
const REMOTE_PROXY_URL = 'https://calm-cell-df4b.gentle-hall-923f.workers.dev';
const GOOGLE_GEMINI_ENDPOINT = `${REMOTE_PROXY_URL}/api/generateContent`;

// ── AI-based analysis using Google Gemini Flash ──
async function analyzeWithAI(imageData, W, H, numQuestions) {
  const base64Image = imageDataToBase64(imageData, W, H);
  const prompt = buildOMRPrompt(numQuestions);

  const inlineRequest = {
    model: GOOGLE_GEMINI_MODEL,
    temperature: 0,
    contents: [
      {
        parts: [
          { text: prompt },
          { inline_data: { mime_type: 'image/jpeg', data: base64Image } }
        ]
      }
    ]
  };

  try {
    return await sendGeminiRequest(inlineRequest, numQuestions);
  } catch (err) {
    console.warn('[analyzeWithAI] inline_data request failed, retrying with fallback prompt:', err.message);
    const fallbackRequest = {
      model: GOOGLE_GEMINI_MODEL,
      temperature: 0,
      contents: [
        {
          parts: [
            {
              text: prompt + `\n\nIMAGE_BASE64:\n${base64Image}`
            }
          ]
        }
      ]
    };
    return await sendGeminiRequest(fallbackRequest, numQuestions);
  }
}

async function sendGeminiRequest(requestBodyObject, numQuestions) {
  const response = await fetch(GOOGLE_GEMINI_ENDPOINT, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(requestBodyObject)
  });

  const responseText = await response.text();
  let data;
  try {
    data = responseText ? JSON.parse(responseText) : {};
  } catch (e) {
    console.error('[sendGeminiRequest] Failed to parse response JSON:', e, 'responseText:', responseText);
    throw new Error(`Invalid JSON response from Gemini proxy: ${e.message}`);
  }

  if (!response.ok) {
    console.error('[sendGeminiRequest] Gemini proxy returned error:', response.status, data);
    throw new Error(`HTTP Error: ${response.status} ${data.error?.message || JSON.stringify(data)}`);
  }

  let aiText = '';
  if (data && data.candidates && Array.isArray(data.candidates) && data.candidates.length > 0) {
    try {
      aiText = data.candidates[0].content.parts[0].text || '';
    } catch (e) {
      console.warn('[sendGeminiRequest] Unable to read standard candidate text:', e);
    }
  }

  if (!aiText) {
    aiText = extractTextFromGoogleResponse(data);
  }

  if (!aiText && typeof data === 'object') {
    aiText = JSON.stringify(data);
  }

  aiText = aiText.replace(/```json/g, '').replace(/```/g, '').trim();

  const jsonMatch = aiText.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    console.error('[sendGeminiRequest] Raw AI response:', aiText);
    throw new Error('لم يتمكن النظام من العثور على JSON في استجابة الـ AI');
  }

  return parseOMRJson(jsonMatch[0], numQuestions);
}

function extractTextFromGoogleResponse(data) {
  if (!data) return '';
  if (typeof data === 'string') return data;
  const parts = [];

  function traverse(node) {
    if (!node || typeof node !== 'object') return;
    if (Array.isArray(node)) return node.forEach(traverse);

    if (typeof node.text === 'string') parts.push(node.text);
    if (typeof node.output_text === 'string') parts.push(node.output_text);

    if (Array.isArray(node.content)) node.content.forEach(traverse);
    if (Array.isArray(node.output)) node.output.forEach(traverse);
    if (Array.isArray(node.candidates)) node.candidates.forEach(traverse);
    if (Array.isArray(node.predictions)) node.predictions.forEach(traverse);
    if (Array.isArray(node.instances)) node.instances.forEach(traverse);
    if (Array.isArray(node.responses)) node.responses.forEach(traverse);

    if (typeof node.response === 'object') traverse(node.response);
    if (typeof node.candidate === 'object') traverse(node.candidate);
  }

  traverse(data);
  return parts.join(' ');
}

function parseOMRJson(text, numQuestions) {
  let raw;
  try {
    raw = JSON.parse(text);
  } catch (e) {
    console.error('[parseOMRJson] Failed to parse JSON:', e);
    console.error('[parseOMRJson] Text was:', text.slice(0, 200));
    throw new Error(`فشل في معالجة استجابة الـ JSON: ${e.message}`);
  }
  
  if (typeof raw !== 'object' || raw === null) {
    throw new Error('استجابة JSON ليست كائناً صالحاً');
  }
  
  const results = {};
  for (let q = 1; q <= numQuestions; q++) {
    const val = raw[String(q)];
    if (val === null || val === undefined) {
      results[q] = null;
      continue;
    }

    if (Array.isArray(val)) {
      const cleaned = val
        .map(item => String(item).trim().toUpperCase())
        .filter(item => ['A','B','C','D'].includes(item));
      const unique = [...new Set(cleaned)];
      if (unique.length === 0) {
        results[q] = null;
      } else if (unique.length === 1) {
        results[q] = unique[0];
      } else {
        results[q] = 'MULTI';
      }
      continue;
    }

    const upper = String(val).trim().toUpperCase();
    if (upper === 'MULTI') {
      results[q] = 'MULTI';
    } else if (['A','B','C','D'].includes(upper)) {
      results[q] = upper;
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

  // 2. تحسين التباين (Contrast Enhancement) قبل Otsu
  const mean = gray.reduce((a, b) => a + b, 0) / gray.length;
  const std = Math.sqrt(gray.reduce((a, b) => a + (b - mean) ** 2, 0) / gray.length);
  for (let i = 0; i < gray.length; i++) {
    gray[i] = Math.max(0, Math.min(255, Math.round((gray[i] - mean) / (std + 1) * 40 + 128)));
  }

  // 3. Otsu threshold
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

  const hLines = findLines(rowDark, H, 0.10, 5);  // تقليل minDensity و minGap
  const vLines = findLines(colDark, W, 0.05, 4);  // أكثر حساسية للخطوط الضعيفة

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
  const MARGIN = 0.15;  // تقليل الهامش لكشف أفضل
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

    // تحسين الخوارزمية: استخدام threshold أقل والحد الأدنى أعلى
    const maxR = Math.max(...ratios);
    const ABS = 0.08;  // حد أدنى أقل (6% تعتم)
    const REL = 0.40;  // نسبة نسبية أقل قليلاً (أكثر تحفظاً)
    const MIN_DIFF = 0.12;  // فرق أدنى بين الإجابة المختارة والثانية
    
    const cands = ratios
      .map((r, i) => ({ label: LABELS[i], ratio: r }))
      .filter(x => x.ratio >= ABS && x.ratio >= maxR * REL);

    if (!cands.length) {
      results[q+1] = null;
    } else if (cands.length === 1) {
      results[q+1] = cands[0].label;
    } else {
      cands.sort((a,b) => b.ratio - a.ratio);
      const gap = cands[0].ratio - cands[1].ratio;
      // إذا كان الفرق أقل من الحد الأدنى = إجابات متعددة
      results[q+1] = gap < MIN_DIFF ? 'MULTI' : cands[0].label;
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
  const model = scannerModuleState.selectedModel || scannerModuleState.selectedSubject?.models[0];
  if (!model) return showToast('يرجى اختيار النموذج أولاً', 'error');

  const numQuestions = model.questions.length;
  const el = document.getElementById('scan-result-content');
  
  console.log('[processOMRImage] Starting OMR processing. Image:', W, 'x', H, 'pixels. Model questions:', numQuestions);

  el.innerHTML = `
    <div style="text-align:center;padding:20px;">
      <div style="font-size:32px;margin-bottom:12px;">🤖</div>
      <div style="color:var(--accent);font-size:15px;font-weight:600;">جارٍ التحليل بالذكاء الاصطناعي...</div>
      <div style="color:var(--text2);font-size:12px;margin-top:6px;">يتم تحليل ورقة الإجابات عبر Google Gemini</div>
      <div style="margin-top:16px;display:flex;justify-content:center;">
        <div style="width:40px;height:4px;background:var(--accent);border-radius:2px;
                    animation:omrPulse 1s ease-in-out infinite alternate;"></div>
      </div>
    </div>
    <style>
      @keyframes omrPulse { from{opacity:.3;transform:scaleX(.5)} to{opacity:1;transform:scaleX(1)} }
    </style>`;

  // Try AI first, fallback to pixel engine
  console.log('[processOMRImage] Calling analyzeWithAI...');
  analyzeWithAI(imageData, W, H, numQuestions)
    .catch(err => {
      console.warn('[processOMRImage] AI analysis failed. Error:', err.message);
      console.warn('[processOMRImage] Full error:', err);
      el.innerHTML = `
        <div style="color:var(--warn);font-size:12px;padding:8px;margin-bottom:8px;
                    background:#ffb84f11;border-radius:6px;border:1px solid #ffb84f33;">
          ⚠️ التحليل الذكي غير متاح (${err.message}) — جارٍ استخدام محرك الكشف التقليدي
        </div>`;
      console.log('[processOMRImage] Falling back to pixel analysis...');
      return analyzeWithPixels(imageData, W, H, numQuestions);
    })
    .then(results => {
      console.log('[processOMRImage] Analysis complete. Results:', results);
      scannerModuleState.currentScannedAnswers = results;
      const pending = [];
      Object.entries(results).forEach(([q, v]) => {
        if (v === null || v === 'MULTI') pending.push(parseInt(q));
      });
      scannerModuleState.pendingManualQuestions = pending;
      console.log('[processOMRImage] Pending manual questions:', pending);
      showScanResults(results, model);
      if (pending.length > 0) showManualOverride(pending, model);
      else computeAndShowScore(results, model);
    })
    .catch(err => {
      console.error('[processOMRImage] FATAL ERROR:', err);
      el.innerHTML = `<div style="color:var(--danger);padding:12px;border-radius:6px;background:#ff4f6a11;">❌ خطأ في تحليل الصورة: ${err.message}</div>`;
    });
}

function showScanResults(e,t){const n=t.questions.length,a=Object.values(e).filter(e=>e&&"MULTI"!==e).length,r=Object.values(e).filter(e=>null===e||"MULTI"===e).length;let s=`<div style="margin-bottom:10px;font-size:12px;color:var(--text2);">\n    تم اكتشاف <strong style="color:var(--accent3);">${a}</strong> إجابة من أصل ${n}\n    ${r>0?`| <strong style="color:var(--warn);">${r}</strong> غير واضحة`:""}\n  </div>`;s+='<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:4px;">';for(let t=1;t<=n;t++){const n=e[t];let a="var(--bg)",r="var(--text2)",o=n||"?";null===n&&(a="#ffb84f22",r="var(--warn)",o="؟"),"MULTI"===n&&(a="#ff4f6a22",r="var(--danger)",o="!!"),n&&"MULTI"!==n&&(a="#4f7cff22",r="var(--accent)"),s+=`<div style="padding:4px;background:${a};border-radius:6px;text-align:center;border:1px solid ${a};">\n      <div style="font-size:9px;color:var(--text3);">${t}</div>\n      <div style="font-size:13px;font-weight:700;color:${r};">${o}</div>\n    </div>`}s+="</div>",document.getElementById("scan-result-content").innerHTML=s}function showManualOverride(e,t){const n=document.getElementById("manual-override-panel");document.getElementById("manual-override-list").innerHTML=e.map(e=>{const n=t.questions[e-1],a=n?n.choices.map(e=>e.label):["A","B","C","D"],r="MULTI"===scannerModuleState.currentScannedAnswers[e];return`<div style="padding:8px;background:var(--bg);border-radius:8px;margin-bottom:8px;border:1px solid var(--border);">\n      <div style="font-size:12px;color:var(--text2);margin-bottom:6px;">\n        <strong style="color:var(--warn);">س${e}:</strong>\n        ${n?n.text.slice(0,60)+(n.text.length>60?"...":""):"سؤال رقم "+e}\n        ${r?'<span style="color:var(--danger);font-size:11px;"> (إجابات متعددة)</span>':'<span style="color:var(--warn);font-size:11px;"> (لم تُكتشف)</span>'}\n      </div>\n      <div style="display:flex;gap:6px;flex-wrap:wrap;">\n        ${a.map(t=>`<button onclick="setManualAnswer(${e},'${t}',this)"\n                  style="padding:6px 14px;border-radius:6px;border:1px solid var(--border2);\n                         background:var(--bg3);color:var(--text);cursor:pointer;\n                         font-family:Cairo,sans-serif;font-weight:600;font-size:13px;\n                         transition:all .2s;"\n                  data-q="${e}" data-label="${t}">${t}</button>`).join("")}\n        <button onclick="setManualAnswer(${e},'SKIP',this)"\n                style="padding:6px 10px;border-radius:6px;border:1px solid #ffb84f44;\n                       background:#ffb84f11;color:var(--warn);cursor:pointer;\n                       font-family:Cairo,sans-serif;font-weight:600;font-size:12px;"\n                data-q="${e}" data-label="SKIP">تجاهل</button>\n      </div>\n    </div>`}).join(""),n.style.display="block"}function setManualAnswer(e,t,n){n.parentElement.querySelectorAll("button").forEach(e=>{e.style.background="var(--bg3)",e.style.color="var(--text)",e.style.borderColor="var(--border2)"}),n.style.background="SKIP"===t?"#ffb84f33":"var(--accent)",n.style.color="SKIP"===t?"var(--warn)":"#fff",n.style.borderColor="SKIP"===t?"var(--warn)":"var(--accent)",scannerModuleState.currentScannedAnswers[e]="SKIP"===t?null:t}function confirmManualAnswers(){if(scannerModuleState.pendingManualQuestions.filter(e=>{const t=scannerModuleState.currentScannedAnswers[e];return null===t||"MULTI"===t}).length>0){const e=scannerModuleState.pendingManualQuestions.filter(e=>"MULTI"===scannerModuleState.currentScannedAnswers[e]);if(e.length>0)return void showToast(`يرجى تحديد إجابة للأسئلة: ${e.join(", ")}`,"error")}const e=scannerModuleState.selectedModel||scannerModuleState.selectedSubject?.models[scannerModuleState.selectedModelIdx||0];e&&(document.getElementById("manual-override-panel").style.display="none",computeAndShowScore(scannerModuleState.currentScannedAnswers,e))}function computeAndShowScore(e,t){let n=0;const a=t.questions.length;t.questions.forEach((t,a)=>{const r=e[a+1];r&&r===t.correctLabel&&n++});const r=document.getElementById("score-summary-panel");document.getElementById("score-display").textContent=`${n} / ${a}`,document.getElementById("score-percent").textContent=`${Math.round(n/a*100)}%`,r.style.display="block",scannerModuleState._pendingResult={studentId:scannerModuleState.currentStudentId,modelName:t.name,modelLetter:t.letter,subject:scannerModuleState.selectedSubject?.name||"",answers:{...e},correct:n,total:a,score:n,percent:Math.round(n/a*100),timestamp:(new Date).toISOString()}}function saveAndNextStudent(){scannerModuleState._pendingResult&&(scannerModuleState.allResults=scannerModuleState.allResults.filter(e=>!(e.studentId===scannerModuleState._pendingResult.studentId&&e.subject===scannerModuleState._pendingResult.subject)),scannerModuleState.allResults.push(scannerModuleState._pendingResult),saveScanResults(scannerModuleState.allResults),renderResultsTable(),showToast(`تم حفظ نتيجة الطالب ${scannerModuleState._pendingResult.studentId} ✓`,"success"),scannerModuleState._pendingResult=null,scannerModuleState.currentStudentId=null,scannerModuleState.currentScannedAnswers=null,scannerModuleState.pendingManualQuestions=[],document.getElementById("scanner-current-student").textContent="-",document.getElementById("scanner-student-id").value="",document.getElementById("scanner-camera-card").style.display="none",document.getElementById("scan-result-content").textContent="اضغط على زر المسح لبدء العملية...",document.getElementById("manual-override-panel").style.display="none",document.getElementById("score-summary-panel").style.display="none",stopCamera(),setTimeout(()=>document.getElementById("scanner-student-id").focus(),100))}function rescanCurrent(){scannerModuleState.currentScannedAnswers=null,scannerModuleState.pendingManualQuestions=[],scannerModuleState._pendingResult=null,document.getElementById("scan-result-content").textContent="اضغط على زر المسح لبدء العملية...",document.getElementById("manual-override-panel").style.display="none",document.getElementById("score-summary-panel").style.display="none",scannerModuleState.cameraStream||startCamera()}function renderResultsTable(){const e=document.getElementById("scanner-results-card"),t=document.getElementById("results-tbody"),n=document.getElementById("results-stats");if(!t)return;const a=scannerModuleState.selectedSubject?.name||"",r=a?scannerModuleState.allResults.filter(e=>e.subject===a):scannerModuleState.allResults;if(0===r.length)return void(e.style.display="none");e.style.display="block",t.innerHTML=r.map((e,t)=>{const n=e.percent??Math.round(e.correct/e.total*100),a=n>=60?"var(--accent3)":n>=50?"var(--warn)":"var(--danger)";return`<tr>\n      <td>${t+1}</td>\n      <td><strong style="color:var(--accent);">${e.studentId}</strong></td>\n      <td>${e.modelName||e.modelLetter||"-"}</td>\n      <td>${e.correct} / ${e.total}</td>\n      <td><strong style="color:${a};">${e.score}</strong></td>\n      <td><span style="color:${a};">${n}%</span></td>\n      <td>\n        <button onclick="deleteResult('${e.studentId}','${e.subject}')"\n                style="padding:4px 10px;background:#ff4f6a22;color:var(--danger);border:none;\n                       border-radius:4px;cursor:pointer;font-family:Cairo,sans-serif;font-size:11px;">حذف</button>\n      </td>\n    </tr>`}).join("");const s=r.length>0?Math.round(r.reduce((e,t)=>e+(t.percent??0),0)/r.length):0,o=r.filter(e=>(e.percent??0)>=60).length;n.innerHTML=`إجمالي الطلاب: <strong>${r.length}</strong> | متوسط الدرجات: <strong style="color:var(--accent);">${s}%</strong> | الناجحون (≥60%): <strong style="color:var(--accent3);">${o}</strong>`}function deleteResult(e,t){confirm(`حذف نتيجة الطالب ${e}؟`)&&(scannerModuleState.allResults=scannerModuleState.allResults.filter(n=>!(n.studentId===e&&n.subject===t)),saveScanResults(scannerModuleState.allResults),renderResultsTable(),showToast("تم حذف النتيجة","success"))}function clearAllResults(){if(!confirm("هل أنت متأكد من حذف جميع النتائج؟ لا يمكن التراجع عن هذا الإجراء."))return;const e=scannerModuleState.selectedSubject?.name||"";scannerModuleState.allResults=e?scannerModuleState.allResults.filter(t=>t.subject!==e):[],saveScanResults(scannerModuleState.allResults),renderResultsTable(),showToast("تم مسح النتائج","success")}function exportResultsExcel(){const e=scannerModuleState.selectedSubject?.name||"",t=e?scannerModuleState.allResults.filter(t=>t.subject===e):scannerModuleState.allResults;if(0===t.length)return void showToast("لا توجد نتائج للتصدير","error");const n=XLSX.utils.book_new(),a=[["الرقم الجامعي","النموذج","المادة","الإجابات الصحيحة","مجموع الأسئلة","العلامة","النسبة المئوية","تاريخ المسح"]];t.forEach(e=>{const t=e.percent??Math.round(e.correct/e.total*100);a.push([e.studentId,e.modelName||e.modelLetter||"",e.subject,e.correct,e.total,e.score,t+"%",e.timestamp?new Date(e.timestamp).toLocaleString("ar-EG"):""])});const r=XLSX.utils.aoa_to_sheet(a);r["!cols"]=[{wch:18},{wch:14},{wch:20},{wch:16},{wch:14},{wch:10},{wch:14},{wch:22}],XLSX.utils.book_append_sheet(n,r,"نتائج الامتحان");const s=`نتائج_${e||"الامتحان"}_${(new Date).toLocaleDateString("ar-EG").replace(/\//g,"-")}.xlsx`;XLSX.writeFile(n,s),showToast("تم تحميل ملف Excel ✓","success")}

