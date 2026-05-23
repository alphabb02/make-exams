let questions=[],models=[],currentModelIdx=0,colMap= {
  q:"",a:"",b:"",c:"",d:"",ans:"",topic:"",diff:""
},opts= {
  shuffleQ:!0,shuffleA:!0,unique:!0,header:!0,student:!0,modelNum:!0,answerKey:!0,answerTable:!0,answerTablePos:"start"
};
function detectDir() {
  const e=questions.slice(0,30).map(e=>(e[colMap.q]||"").trim()).join(" ");
  return(e.match(/[\u0600-\u06FF]/g)||[]).length>=(e.match(/[A-Za-z]/g)||[]).length?"rtl":"ltr"
}
let design= {
  headerBg:"#1e3a8a",headerText:"#ffffff",fontSize:14,layoutCols:2,answerStyle:"normal"
},examCfg= {

};
const STORAGE_KEY="exam_projects_v1",ENCRYPTION_KEY="exam_builder_2026";
function encryptProject(e) {
  try {
    const t=JSON.stringify(e),n=btoa(unescape(encodeURIComponent(t)));
    let o="";
    for(let e=0; e<n.length; e++)o+=String.fromCharCode(n.charCodeAt(e)^ENCRYPTION_KEY.charCodeAt(e%17));
    return btoa(o)
  }
  catch(e) {
    return console.error("Encryption error:",e),null
  }
}
function decryptProject(e) {
  try {
    const t=atob(e);
    let n="";
    for(let e=0; e<t.length; e++)n+=String.fromCharCode(t.charCodeAt(e)^ENCRYPTION_KEY.charCodeAt(e%17));
    const o=decodeURIComponent(escape(atob(n)));
    return JSON.parse(o)
  }
  catch(e) {
    return console.error("Decryption error:",e),null
  }
}
function getAllProjects() {
  try {
    const e=localStorage.getItem(STORAGE_KEY);
    return e?JSON.parse(e):[]
  }
  catch(e) {
    return console.error("Get projects error:",e),[]
  }
}
function saveProjectToStorage(e) {
  try {
    const t=getAllProjects(),n=encryptProject(e);
    if(!n)return console.warn("Encryption failed"),null;
    const o=e.id||Date.now(),r=e.examCfg?.subject||"مشروع بدون اسم",a= {
      id:o,name:r,questions:e.questions?.length||0,models:e.models?.length||0,date:(new Date).toISOString(),encrypted:n
    },s=t.findIndex(e=>e.id===o);
    s>=0?t[s]=a:t.push(a);
    const l=JSON.stringify(t),i=new Blob([l]).size/1024;
    return i>5e3&&console.warn(`Storage size: ${i.toFixed(0)}KB - approaching limit`),localStorage.setItem(STORAGE_KEY,l),console.log(`Project saved: "${r}" — ${a.questions} questions, ${a.models} models`),a.id
  }
  catch(e) {
    return console.error("Save error:",e),null
  }
}
function loadProjectFromStorage(e) {
  try {
    const t=getAllProjects().find(t=>t.id===e);
    return t&&t.encrypted?decryptProject(t.encrypted):null
  }
  catch(e) {
    return console.error("Load error:",e),null
  }
}
function deleteProjectFromStorage(e) {
  try {
    let t=getAllProjects();
    return t=t.filter(t=>t.id!==e),localStorage.setItem(STORAGE_KEY,JSON.stringify(t)),!0
  }
  catch(e) {
    return console.error("Delete error:",e),!1
  }
}
function autoSaveCurrentProject() {
  if(0===questions.length)return;
  window.currentProjectId||(window.currentProjectId=Date.now());
  const e=document.getElementById("subjectName")?.value?.trim()||examCfg.subject||"";
  e&&(examCfg.subject=e);
  const t= {
    id:window.currentProjectId,questions:questions,colMap:colMap,opts:opts,design:design,examCfg:examCfg,models:models
  };
  window.currentProjectId=saveProjectToStorage(t)
}
function hasMathJax() {
  return"undefined"!=typeof window&&window.MathJax&&window.MathJax.typesetPromise
}
function hasMath(e) {
  return!!e&&(/\$[\s\S]+?\$/.test(e)||/\\\([\s\S]+?\\\)/.test(e)||/\\\[[\s\S]+?\\\]/.test(e)||/\\(?:frac|sqrt|sum|int|lim|vec|hat|bar|over|alpha|beta|gamma|delta|theta|pi|sigma|omega|lambda|mu|infty|cdot|times|div|pm|leq|geq|neq|approx|rightarrow|leftarrow|Rightarrow|forall|exists|in|subset|cup|cap|mathbb|mathbf|mathrm)\b/.test(e)||/[²³¹⁰⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿ]/.test(e)||/[₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎]/.test(e)||/[∑∏∫∂∇∞∈∉⊂⊃∪∩≤≥≠≈±×÷√∝∀∃∄⊕⊗∧∨¬→←↔⇒⇔αβγδεζηθικλμνξπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΠΡΣΤΥΦΧΨΩ]/.test(e)||/\b(?:sqrt|log|ln|sin|cos|tan|cot|sec|csc|lim|max|min)\s*[\(\d]/.test(e)||/\d+\s*[\^]\s*\d+/.test(e)||/\d+\s*\/\s*\d+/.test(e))
}
function normalizeToLatex(e) {
  if(!e)return e;
  let t=e;
  if(/\$|\\[(\[]/.test(t))return t;
  const n= {
    "²":"^{2}","³":"^{3}","¹":"^{1}","⁰":"^{0}","⁴":"^{4}","⁵":"^{5}","⁶":"^{6}","⁷":"^{7}","⁸":"^{8}","⁹":"^{9}","ⁿ":"^{n}","⁺":"^{+}","⁻":"^{-}"
  },o= {
    "₀":"_{0}","₁":"_{1}","₂":"_{2}","₃":"_{3}","₄":"_{4}","₅":"_{5}","₆":"_{6}","₇":"_{7}","₈":"_{8}","₉":"_{9}"
  };
  let r=!1,a=!1;
  for(const[e,o]of Object.entries(n))t.includes(e)&&(t=t.replaceAll(e,o),r=!0);
  for(const[e,n]of Object.entries(o))t.includes(e)&&(t=t.replaceAll(e,n),a=!0);
  if(/[a-zA-Z0-9]\^[ {
    0-9a-zA-Z]/.test(t)&&(r=!0),t=t.replace(/\bsqrt\s*\(([^)]+)\)/g,"\\sqrt{$1}"),t=t.replace(/\bsqrt\s+(\S+)/g,"\\sqrt{$1}"),t=t.replace(/√\s*\(([^)]+)\)/g,"\\sqrt{$1}"),t=t.replace(/√\s*(\S+)/g,"\\sqrt{$1}"),t=t.replace(/\b(\d+)\s*\/\s*(\d+)\b/g,"\\frac{$1}{$2}"),(r||a||/\\(?:sqrt|frac|sum|int|alpha|beta|gamma|delta|theta|pi|sigma|omega|lambda|sin|cos|tan|log|ln|lim)/.test(t)||/[∑∏∫∂∇∞∈∉⊂⊃∪∩≤≥≠≈±×÷√∝⊕⊗∧∨→←↔⇒⇔]/.test(t))&&!/\$/.test(t)) {
      const e=t.trim().split(/\s+/).length,n=(t.match(/[\\ {

      }
      \^_]|\\[a-zA-Z]+/g)||[]).length;
      t=n>0&&(n/e>.3||e<=5)?`$${t}$`:t.replace(/((?:[a-zA-Z0-9][\^_][ {

      }
      \d\w]*|\\[a-zA-Z]+(?:\ {
        [^
      }
      ]*\
    })*)+)/g,"$$$1$$$")
  }
  return t
}
function renderMath(e) {
  return e?/<[a-zA-Z]/.test(e)?e.replace(/(?<=>|^)([^<]+)(?=<|$)/g,e=>renderMath(e)):hasMath(e)?normalizeToLatex(e):e:""
}
function typesetMath(e) {
  hasMathJax()&&MathJax.typesetPromise([e]).catch(e=>console.warn("MathJax:",e))
}
function mathJaxScript() {
  return"\n  <script>\n    window.MathJax = {\n      tex: {\n        inlineMath: [['$','$'], ['\\\\(','\\\\)']],\n        displayMath: [['$$','$$'], ['\\\\[','\\\\]']],\n        packages: {'[+]': ['ams','boldsymbol']},\n        tags: 'none'\n      },\n      options: { skipHtmlTags: ['script','noscript','style','textarea','pre'] },\n      startup: { typeset: true }\n    };\n  <\/script>\n  <script src=\"https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js\" id=\"MathJax-script\" async><\/script>"
}
let bankFiles=[],multiFileMode=!1;
function toggleMultiFileMode(e) {
  multiFileMode=e;
  const t=document.getElementById("multi-file-container");
  if(!t)return;
  t.style.display=e?"block":"none";
  const n=document.getElementById("dropZone");
  n&&(n.style.display=e?"none":"block")
}
function addFileToBank(e) {
  if(!e)return;
  const t=new FileReader;
  t.onload=t=> {
    try {
      const n=XLSX.read(t.target.result, {
        type:"array"
      }),o=n.Sheets[n.SheetNames[0]],r=XLSX.utils.sheet_to_json(o, {
        header:1
      });
      if(r.length<2)return void showToast("الملف فارغ","error");
      const a=r[0].map(e=>String(e||"").trim()),s=r.slice(1).filter(e=>e.some(e=>void 0!==e&&""!==e)).map(e=> {
        const t= {

        };
        return a.forEach((n,o)=>t[n]=void 0!==e[o]?String(e[o]).trim():""),t
      });
      bankFiles.push( {
        name:e.name.replace(/\.[^/.]+$/,""),questions:s,alloc:s.length,headers:a
      }),renderBankFilesList(),showToast(`✓ تم إضافة "${e.name}" (${s.length} سؤال)`,"success")
    }
    catch(e) {
      showToast("خطأ في قراءة الملف: "+e.message,"error")
    }
  },t.readAsArrayBuffer(e)
}
function removeBankFile(e) {
  bankFiles.splice(e,1),renderBankFilesList()
}
function updateBankFileAlloc(e,t) {
  const n=parseInt(t)||0;
  bankFiles[e]&&(bankFiles[e].alloc=Math.min(n,bankFiles[e].questions.length)),renderBankFilesList()
}
function renderBankFilesList() {
  const e=document.getElementById("bank-files-list");
  if(!e)return;
  const t=bankFiles.reduce((e,t)=>e+t.alloc,0),n=parseInt(document.getElementById("qPerModel")?.value||20),o=t>=n?"✓":"⚠️",r=t>=n?"var(--accent3)":"var(--warn)";
  e.innerHTML=`\n    <div style="margin-bottom:12px;padding:10px;background:var(--bg3);border-radius:8px;">\n      <div style="font-size:13px;color:var(--text2);margin-bottom:6px;">\n        <span style="color:${r};font-weight:700;">${o} إجمالي المخصص: ${t} / ${n}</span>\n      </div>\n      ${bankFiles.map((e,t)=>`\n        <div style="background:var(--card);border:1px solid var(--border);border-radius:6px;\n                    padding:10px;margin-bottom:8px;display:grid;\n                    grid-template-columns:1fr 80px 50px;gap:8px;align-items:center;">\n          <div>\n            <div style="font-size:13px;font-weight:600;color:var(--text);">$ {
    e.name
  }
  </div>\n            <div style="font-size:11px;color:var(--text3);">$ {
    e.questions.length
  }
  سؤال متاح</div>\n          </div>\n          <input type="number" min="0" max="${e.questions.length}" value="${e.alloc}"\n                 onchange="updateBankFileAlloc(${t}, this.value)"\n                 style="padding:6px;border:1px solid var(--border);border-radius:4px;\n                        background:var(--bg3);color:var(--text);font-family:inherit;font-size:13px;">\n          <button onclick="removeBankFile(${t})"\n                  style="padding:6px 10px;background:var(--danger);color:#fff;border:none;\n                         border-radius:4px;font-family:inherit;cursor:pointer;font-size:11px;">✕</button>\n        </div>\n      `).join("")}\n    </div>`
}
function mergeQuestionsFromBank() {
  if(!multiFileMode||0===bankFiles.length)return questions;
  const e=[];
  bankFiles.forEach(t=> {
    const n=shuffle(t.questions).slice(0,t.alloc);
    e.push(...n)
  });
  const t=parseInt(document.getElementById("qPerModel")?.value||20)*parseInt(document.getElementById("numModels")?.value||4);
  if(e.length<t&&bankFiles.length>0) {
    const n=bankFiles[bankFiles.length-1].questions.filter(t=>!e.some(e=>e[colMap.q]===t[colMap.q])),o=Math.min(n.length,t-e.length);
    e.push(...n.slice(0,o))
  }
  return e
}
let multiPartEnabled=!1,mpParts=[];
const MP_COLORS=["#4f7cff","#7c5cfc","#00d4aa","#ffb84f","#ff4f6a","#06b6d4","#a855f7","#f97316"];
function toggleMultiPartMode(e) {
  e.classList.toggle("on"),multiPartEnabled=e.classList.contains("on");
  const t=document.getElementById("multipart-card");
  t&&(t.style.display=multiPartEnabled?"block":"none"),multiPartEnabled&&renderMpPartsList()
}
function handleMpFile(e) {
  const t=e.target.files[0];
  if(!t)return;
  const n=new FileReader;
  n.onload=e=> {
    try {
      const n=XLSX.read(e.target.result, {
        type:"array"
      }),o=n.Sheets[n.SheetNames[0]],r=XLSX.utils.sheet_to_json(o, {
        header:1
      });
      if(r.length<2)return void showToast("الملف فارغ","error");
      const a=r[0].map(e=>String(e||"").trim()),s= {
        ...colMap
      };
      autoMap(a);
      const l= {
        ...colMap
      };
      colMap=s;
      const i=r.slice(1).filter(e=>e.some(e=>void 0!==e&&""!==e)).map(e=> {
        const t= {

        };
        return a.forEach((n,o)=>t[n]=void 0!==e[o]?String(e[o]).trim():""),t.__colMap=l,t
      }),d=t.name.replace(/\.[^.]+$/,""),c=MP_COLORS[mpParts.length%MP_COLORS.length],p=parseInt(document.getElementById("qPerModel")?.value||20),u=Math.max(1,Math.floor(p/(mpParts.length+1)));
      mpParts.push( {
        name:d,questions:i,requested:Math.min(u,i.length),color:c,colMap:l
      }),document.getElementById("mpFileInput").value="",renderMpPartsList(),showToast(`✅ تم إضافة "${d}" — ${i.length} سؤال`,"success")
    }
    catch(e) {
      showToast("خطأ في قراءة الملف: "+e.message,"error")
    }
  },n.readAsArrayBuffer(t)
}
function removeMpPart(e) {
  mpParts.splice(e,1),renderMpPartsList()
}
function updateMpPartCount(e,t) {
  mpParts[e].requested=Math.min(parseInt(t)||0,mpParts[e].questions.length),renderMpDistSummary()
}
function renderMpPartsList() {
  const e=document.getElementById("mp-parts-list"),t=document.getElementById("mp-dist-summary");
  if(!mpParts.length)return e.innerHTML="",void(t&&(t.style.display="none"));
  e.innerHTML=mpParts.map((e,t)=>`\n    <div style="display:flex;align-items:center;gap:10px;padding:10px 12px;margin-bottom:8px;\n                background:var(--bg3);border-radius:var(--radius-sm);\n                border:1px solid var(--border);border-right:3px solid ${e.color};">\n      <div style="flex:1;min-width:0;">\n        <div style="font-size:13px;font-weight:600;color:var(--text);white-space:nowrap;\n                    overflow:hidden;text-overflow:ellipsis;">${e.name}</div>\n        <div style="font-size:11px;color:var(--text3);margin-top:2px;">\n          متاح: <strong style="color:${e.color};">${e.questions.length}</strong> سؤال\n        </div>\n      </div>\n      <div style="display:flex;align-items:center;gap:8px;flex-shrink:0;">\n        <label style="font-size:11px;color:var(--text2);white-space:nowrap;">عدد الأسئلة:</label>\n        <input type="number" min="0" max="${e.questions.length}"\n               value="${e.requested||0}"\n               onchange="updateMpPartCount(${t}, this.value)"\n               style="width:70px;text-align:center;padding:6px;border-radius:6px;\n                      border:1px solid ${e.color};background:var(--bg);\n                      color:var(--text);font-family:'Cairo',sans-serif;font-size:13px;">\n      </div>\n      <button onclick="removeMpPart(${t})"\n              style="width:28px;height:28px;border-radius:6px;border:none;\n                     background:#ff4f6a22;color:var(--danger);cursor:pointer;\n                     font-size:14px;display:flex;align-items:center;justify-content:center;flex-shrink:0;">✕</button>\n    </div>`).join(""),t&&(t.style.display="block"),renderMpDistSummary();
  const n=document.getElementById("mpDropZone");
  n&&!n._mpDragSetup&&(n._mpDragSetup=!0,n.addEventListener("dragover",e=> {
    e.preventDefault(),n.classList.add("drag")
  }),n.addEventListener("dragleave",()=>n.classList.remove("drag")),n.addEventListener("drop",e=> {
    e.preventDefault(),n.classList.remove("drag"),handleMpFile( {
      target: {
        files:e.dataTransfer.files
      }
    })
  }))
}
function renderMpDistSummary() {
  const e=mpParts.reduce((e,t)=>e+(t.requested||0),0),t=parseInt(document.getElementById("qPerModel")?.value||20),n=document.getElementById("mp-dist-bars"),o=document.getElementById("mp-dist-warning");
  n&&(n.innerHTML=mpParts.map(t=> {
    const n=e>0?Math.round(t.requested/e*100):0;
    return`\n      <div style="margin-bottom:6px;">\n        <div style="display:flex;justify-content:space-between;font-size:11px;\n                    color:var(--text2);margin-bottom:3px;">\n          <span>${t.name}</span>\n          <span style="color:${t.color};font-weight:600;">${t.requested} سؤال (${n}%)</span>\n        </div>\n        <div style="height:6px;background:var(--border);border-radius:3px;overflow:hidden;">\n          <div style="height:100%;width:${n}%;background:${t.color};border-radius:3px;\n                      transition:width .4s;"></div>\n        </div>\n      </div>`
  }).join("")+`\n    <div style="font-size:12px;color:var(--text2);margin-top:8px;padding-top:8px;\n                border-top:1px solid var(--border);display:flex;justify-content:space-between;">\n      <span>الإجمالي المخصص:</span>\n      <strong style="color:${e===t?"var(--accent3)":e>t?"var(--danger)":"var(--warn)"};">\n        ${e} / ${t}\n      </strong>\n    </div>`,o&&(e>t?(o.style.display="block",o.style.background="#ff4f6a11",o.style.border="1px solid #ff4f6a44",o.style.color="var(--danger)",o.textContent=`⚠️ الإجمالي (${e}) يتجاوز عدد أسئلة النموذج (${t}). سيتم الاقتصار على ${t}.`):e<t&&e>0&&mpParts.length>0?(o.style.display="block",o.style.background="#ffb84f11",o.style.border="1px solid #ffb84f44",o.style.color="var(--warn)",o.textContent=`ℹ️ النقص (${t-e} سؤال) سيُكمَّل تلقائياً من آخر ملف مرفوع: "${mpParts[mpParts.length-1].name}".`):o.style.display="none"))
}
function buildMpQuestionsPool() {
  if(!multiPartEnabled||0===mpParts.length)return null;
  const e=parseInt(document.getElementById("qPerModel")?.value||20),t=[];
  for(let e=0; e<mpParts.length; e++) {
    const n=mpParts[e];
    let o=n.requested||0;
    o=Math.min(o,n.questions.length);
    const r=shuffle([...n.questions]).slice(0,o).map(e=>( {
      ...e,__partName:n.name,__partColor:n.color,__partColMap:n.colMap
    }));
    t.push(...r)
  }
  if(t.length<e&&mpParts.length>0) {
    const n=mpParts[mpParts.length-1],o=new Set(t.map(e=>e[n.colMap.q||Object.keys(e)[0]]||"")),r=n.questions.filter(e=> {
      const t=n.colMap.q||Object.keys(e)[0];
      return!o.has(e[t]||"")
    }),a=Math.min(r.length,e-t.length),s=shuffle(r).slice(0,a).map(e=>( {
      ...e,__partName:n.name,__partColor:n.color,__partColMap:n.colMap
    }));
    t.push(...s)
  }
  return mpParts.length>0&&(colMap= {
    ...mpParts[0].colMap
  }),t
}
function goPanel(e) {
  if(e>0&&e<4&&0===questions.length&&!multiPartEnabled)showToast("يرجى استيراد الأسئلة أولاً","error");
  else if(e>0&&e<4&&multiPartEnabled&&0===mpParts.length)showToast("يرجى إضافة أجزاء المادة في إعدادات النماذج","error");
  else {
    if(document.querySelectorAll(".panel").forEach((t,n)=>t.classList.toggle("active",n===e)),document.querySelectorAll(".step-btn").forEach((t,n)=>t.classList.toggle("active",n===e)),4===e&&initScannerPanel(),2===e) {
      const e=document.getElementById("answerStyle");
      e&&(e.value=design.answerStyle||"normal")
    }
    updatePreview()
  }
}
function toggleOpt(e) {
  e.classList.toggle("on");
  const t=e.id.replace("tgl-",""),n= {
    "shuffle-q":"shuffleQ","shuffle-a":"shuffleA",unique:"unique",header:"header",student:"student",modelnum:"modelNum",answerkey:"answerKey",answertable:"answerTable"
  };
  void 0!==n[t]&&(opts[n[t]]=e.classList.contains("on")),autoSaveCurrentProject()
}
function setAnswerTablePos(e) {
  opts.answerTablePos=e;
  const t=document.getElementById("pos-start"),n=document.getElementById("pos-end");
  t&&n&&("start"===e?(t.style.cssText="padding:5px 10px;border-radius:6px;font-family:Cairo,sans-serif;font-size:11px;font-weight:600;cursor:pointer;border:1px solid var(--accent);background:var(--accent);color:#fff;",n.style.cssText="padding:5px 10px;border-radius:6px;font-family:Cairo,sans-serif;font-size:11px;font-weight:600;cursor:pointer;border:1px solid var(--border2);background:transparent;color:var(--text2);"):(n.style.cssText="padding:5px 10px;border-radius:6px;font-family:Cairo,sans-serif;font-size:11px;font-weight:600;cursor:pointer;border:1px solid var(--accent);background:var(--accent);color:#fff;",t.style.cssText="padding:5px 10px;border-radius:6px;font-family:Cairo,sans-serif;font-size:11px;font-weight:600;cursor:pointer;border:1px solid var(--border2);background:transparent;color:var(--text2);"),autoSaveCurrentProject())
}
function initDropZone() {
  const e=document.getElementById("dropZone");
  if(!e)return;
  e.addEventListener("dragover",t=> {
    t.preventDefault(),e.classList.add("drag")
  }),e.addEventListener("dragleave",()=>e.classList.remove("drag")),e.addEventListener("drop",t=> {
    t.preventDefault(),e.classList.remove("drag"),handleFile( {
      target: {
        files:t.dataTransfer.files
      }
    })
  });
  const t=document.getElementById("multiDropZone");
  t&&(t.addEventListener("dragover",e=> {
    e.preventDefault(),t.classList.add("drag")
  }),t.addEventListener("dragleave",()=>t.classList.remove("drag")),t.addEventListener("drop",e=> {
    e.preventDefault(),t.classList.remove("drag"),handleMultiFile( {
      target: {
        files:e.dataTransfer.files
      }
    })
  }))
}
let importMode="single",parts=[];
const PART_COLORS=["#4f7cff","#7c5cfc","#00d4aa","#ffb84f","#ff4f6a","#06b6d4","#a855f7","#f97316"];
function setImportMode(e) {
  importMode=e;
  const t=document.getElementById("mode-single-btn"),n=document.getElementById("mode-multi-btn"),o=document.getElementById("single-mode-area"),r=document.getElementById("multi-mode-area");
  "single"===e?(t.style.border="2px solid var(--accent)",t.style.background="var(--accent)11",t.querySelector("div:nth-child(2)").style.color="var(--accent)",n.style.border="2px solid var(--border)",n.style.background="transparent",n.querySelector("div:nth-child(2)").style.color="var(--text2)",o.style.display="block",r.style.display="none"):(n.style.border="2px solid var(--accent)",n.style.background="var(--accent)11",n.querySelector("div:nth-child(2)").style.color="var(--accent)",t.style.border="2px solid var(--border)",t.style.background="transparent",t.querySelector("div:nth-child(2)").style.color="var(--text2)",o.style.display="none",r.style.display="block")
}
function resetImport() {
  questions=[],parts=[],document.getElementById("preview-table").style.display="none",document.getElementById("stat-total").textContent="0",document.getElementById("stat-topics").textContent="0",document.getElementById("stat-cols").textContent="0",document.getElementById("stat-ready").textContent="-",document.getElementById("sb-total").textContent="0",document.getElementById("sb-topics").textContent="0",renderPartsList()
}
function handleMultiFile(e) {
  const t=e.target.files[0];
  if(!t)return;
  const n=new FileReader;
  n.onload=e=> {
    try {
      const n=XLSX.read(e.target.result, {
        type:"array"
      }),o=n.Sheets[n.SheetNames[0]],r=XLSX.utils.sheet_to_json(o, {
        header:1
      });
      if(r.length<2)return void showToast("الملف فارغ","error");
      const a=r[0].map(e=>String(e||"").trim()),s= {
        ...colMap
      };
      autoMap(a);
      const l= {
        ...colMap
      };
      colMap=s;
      const i=r.slice(1).filter(e=>e.some(e=>void 0!==e&&""!==e)).map(e=> {
        const t= {

        };
        return a.forEach((n,o)=>t[n]=void 0!==e[o]?String(e[o]).trim():""),t.__colMap=l,t
      }),d=t.name.replace(/\.[^.]+$/,""),c=PART_COLORS[parts.length%PART_COLORS.length];
      parts.push( {
        name:d,questions:i,requested:0,color:c,colMap:l
      }),document.getElementById("multiFileInput").value="",renderPartsList(),showToast(`✅ تم إضافة "${d}" — ${i.length} سؤال`,"success")
    }
    catch(e) {
      showToast("خطأ في قراءة الملف: "+e.message,"error")
    }
  },n.readAsArrayBuffer(t)
}
function renderPartsList() {
  const e=document.getElementById("parts-list"),t=document.getElementById("dist-summary");
  if(!parts.length)return e.innerHTML="",void(t.style.display="none");
  e.innerHTML=parts.map((e,t)=>`\n    <div style="display:flex;align-items:center;gap:10px;padding:10px 12px;margin-bottom:8px;\n                background:var(--bg3);border-radius:var(--radius-sm);\n                border:1px solid var(--border);border-right:3px solid ${e.color};">\n      <div style="flex:1;min-width:0;">\n        <div style="font-size:13px;font-weight:600;color:var(--text);white-space:nowrap;\n                    overflow:hidden;text-overflow:ellipsis;">${e.name}</div>\n        <div style="font-size:11px;color:var(--text3);margin-top:2px;">\n          متاح: <strong style="color:${e.color};">${e.questions.length}</strong> سؤال\n        </div>\n      </div>\n      <div style="display:flex;align-items:center;gap:8px;flex-shrink:0;">\n        <label style="font-size:11px;color:var(--text2);white-space:nowrap;">عدد الأسئلة:</label>\n        <input type="number" min="0" max="${e.questions.length}"\n               value="${e.requested||0}"\n               onchange="updatePartCount(${t}, this.value)"\n               style="width:70px;text-align:center;padding:6px;border-radius:6px;\n                      border:1px solid ${e.color};background:var(--bg);\n                      color:var(--text);font-family:'Cairo',sans-serif;font-size:13px;">\n      </div>\n      <button onclick="removePart(${t})"\n              style="width:28px;height:28px;border-radius:6px;border:none;\n                     background:#ff4f6a22;color:var(--danger);cursor:pointer;\n                     font-size:14px;display:flex;align-items:center;justify-content:center;flex-shrink:0;">✕</button>\n    </div>`).join(""),t.style.display="block",renderDistSummary()
}
function updatePartCount(e,t) {
  parts[e].requested=Math.min(parseInt(t)||0,parts[e].questions.length),renderDistSummary()
}
function removePart(e) {
  parts.splice(e,1),renderPartsList()
}
function renderDistSummary() {
  const e=parts.reduce((e,t)=>e+(t.requested||0),0),t=parseInt(document.getElementById("qPerModel")?.value||20),n=document.getElementById("dist-bars"),o=document.getElementById("dist-warning");
  n.innerHTML=parts.map(t=> {
    const n=e>0?Math.round(t.requested/e*100):0;
    return`\n      <div style="margin-bottom:6px;">\n        <div style="display:flex;justify-content:space-between;font-size:11px;\n                    color:var(--text2);margin-bottom:3px;">\n          <span>${t.name}</span>\n          <span style="color:${t.color};font-weight:600;">${t.requested} سؤال (${n}%)</span>\n        </div>\n        <div style="height:6px;background:var(--border);border-radius:3px;overflow:hidden;">\n          <div style="height:100%;width:${n}%;background:${t.color};border-radius:3px;\n                      transition:width .4s;"></div>\n        </div>\n      </div>`
  }).join("")+`\n    <div style="font-size:12px;color:var(--text2);margin-top:8px;padding-top:8px;\n                border-top:1px solid var(--border);display:flex;justify-content:space-between;">\n      <span>الإجمالي المطلوب:</span>\n      <strong style="color:${e===t?"var(--accent3)":e>t?"var(--danger)":"var(--warn)"};">\n        ${e} / ${t}\n      </strong>\n    </div>`,e>t?(o.style.display="block",o.textContent=`⚠️ الإجمالي (${e}) يتجاوز عدد أسئلة النموذج (${t}). سيتم الاقتصار على ${t}.`):e<t&&e>0?(o.style.display="block",o.style.background="#ffb84f11",o.style.borderColor="#ffb84f44",o.style.color="var(--warn)",o.textContent=`ℹ️ النقص (${t-e} سؤال) سيُكمَّل تلقائياً من آخر جزء مرفوع.`):o.style.display="none",e>0&&mergeAndPreviewMulti()
}
function mergeAndPreviewMulti() {
  const e=[],t=parseInt(document.getElementById("qPerModel")?.value||20);
  let n=t;
  for(let o=0; o<parts.length; o++) {
    const r=parts[o];
    let a=r.requested||0;
    o===parts.length-1&&e.length<t&&(a=Math.min(t-e.length,r.questions.length)),a=Math.min(a,r.questions.length,n);
    const s=shuffle([...r.questions]).slice(0,a).map(e=>( {
      ...e,__partName:r.name,__partColor:r.color,__partColMap:r.colMap
    }));
    if(e.push(...s),n-=a,n<=0)break
  }
  parts.length>0&&(colMap= {
    ...parts[0].colMap
  }),questions=e,document.getElementById("stat-total").textContent=questions.length,document.getElementById("stat-topics").textContent=parts.length,document.getElementById("stat-cols").textContent="—",document.getElementById("stat-ready").textContent="✅",document.getElementById("sb-total").textContent=questions.length,document.getElementById("sb-topics").textContent=parts.length,document.getElementById("sb-avg").textContent="4",renderMultiPreview()
}
function renderMultiPreview() {
  document.getElementById("preview-table").style.display="block",document.getElementById("table-head").innerHTML="<th>#</th><th>الجزء</th><th>السؤال</th><th>A</th><th>Answer</th>",document.getElementById("table-body").innerHTML=questions.slice(0,10).map((e,t)=> {
    const n=e.__partColMap||colMap;
    return`<tr>\n        <td>${t+1}</td>\n        <td><span style="background:${e.__partColor||"var(--accent)"}22;\n                         color:${e.__partColor||"var(--accent)"};\n                         padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600;">\n          ${e.__partName||"—"}\n        </span></td>\n        <td>${String(e[n.q]||"").slice(0,50)}</td>\n        <td>${String(e[n.a]||"").slice(0,30)}</td>\n        <td>${String(e[n.ans]||"").slice(0,20)}</td>\n      </tr>`
  }).join("")+(questions.length>10?`<tr><td colspan="5" style="text-align:center;color:var(--text3)">... و ${questions.length-10} سؤال آخر</td></tr>`:"")
}
function handleFile(e) {
  const t=e.target.files[0];
  if(!t)return;
  const n=new FileReader;
  n.onload=e=> {
    try {
      const t=XLSX.read(e.target.result, {
        type:"array"
      }),n=t.Sheets[t.SheetNames[0]];
      processData(XLSX.utils.sheet_to_json(n, {
        header:1
      }))
    }
    catch(e) {
      showToast("خطأ في قراءة الملف: "+e.message,"error")
    }
  },n.readAsArrayBuffer(t)
}
function processData(e) {
  if(e.length<2)return void showToast("الملف فارغ","error");
  const t=e[0].map(e=>String(e||"").trim());
  autoMap(t),questions=e.slice(1).filter(e=>e.some(e=>void 0!==e&&""!==e)).map(e=> {
    const n= {

    };
    return t.forEach((t,o)=>n[t]=void 0!==e[o]?String(e[o]).trim():""),n
  }),updateStats(t),renderPreviewTable(t),autoSaveCurrentProject(),showToast("تم استيراد "+questions.length+" سؤال","success")
}
function autoMap(e) {
  const t=e.map(e=>String(e||"").trim().toLowerCase()),n=n=> {
    for(const o of n) {
      const n=t.findIndex(e=>e===o.toLowerCase());
      if(-1!==n)return e[n]
    }
    return""
  },o=n=> {
    for(const o of n) {
      const n=t.findIndex(e=>e.includes(o.toLowerCase()));
      if(-1!==n)return e[n]
    }
    return""
  };
  colMap.q=n(["Question","question","سؤال","نص السؤال"])||o(["question","سؤال","نص"]),colMap.a=n(["A","a","أ","ا","Choice A","Option A"])||o(["choice a","option a","خيار أ"]),colMap.b=n(["B","b","ب","Choice B","Option B"])||o(["choice b","option b","خيار ب"]),colMap.c=n(["C","c","ج","Choice C","Option C"])||o(["choice c","option c","خيار ج"]),colMap.d=n(["D","d","د","Choice D","Option D"])||o(["choice d","option d","خيار د"]),colMap.ans=n(["Answer","answer","Correct","correct","الإجابة","إجابة","الإجابة الصحيحة"])||o(["answer","correct","إجابة","صحيح"]),colMap.topic=n(["Topic","topic","Unit","unit","محور","وحدة","فصل"])||o(["topic","unit","محور","وحدة"]),colMap.diff=n(["Difficulty","difficulty","Level","level","صعوبة"])||o(["difficulty","level","صعوبة"]),!colMap.q&&e[0]&&(colMap.q=e[0]),!colMap.a&&e[1]&&(colMap.a=e[1]),!colMap.b&&e[2]&&(colMap.b=e[2]),!colMap.c&&e[3]&&(colMap.c=e[3]),!colMap.d&&e[4]&&(colMap.d=e[4]),!colMap.ans&&e[5]&&(colMap.ans=e[5])
}
function updateStats(e) {
  const t=new Set(questions.map(e=>e[colMap.topic]).filter(Boolean));
  document.getElementById("stat-total").textContent=questions.length,document.getElementById("stat-topics").textContent=t.size||"—",document.getElementById("stat-cols").textContent=e.length,document.getElementById("stat-ready").textContent="✅",document.getElementById("sb-total").textContent=questions.length,document.getElementById("sb-topics").textContent=t.size||0,document.getElementById("sb-avg").textContent=[colMap.a,colMap.b,colMap.c,colMap.d].filter(Boolean).length
}
function updateSidebar() {
  const e=new Set(questions.map(e=>e[colMap.topic]).filter(Boolean)),t=document.getElementById("sb-total"),n=document.getElementById("sb-topics"),o=document.getElementById("sb-avg");
  t&&(t.textContent=questions.length),n&&(n.textContent=e.size||0),o&&(o.textContent=[colMap.a,colMap.b,colMap.c,colMap.d].filter(Boolean).length)
}
function renderPreviewTable(e) {
  const t=e.slice(0,6);
  document.getElementById("table-head").innerHTML="<th>#</th>"+t.map(e=>`<th>${e}</th>`).join(""),document.getElementById("table-body").innerHTML=questions.slice(0,8).map((e,n)=>`<tr><td>${n+1}</td>${t.map(t=>`<td>$ {
    String(e[t]||"").slice(0,60)
  }
  </td>`).join("")}</tr>`).join("")+(questions.length>8?`<tr><td colspan="${t.length+1}" style="text-align:center;color:var(--text3)">\n           ... و ${questions.length-8} سؤال آخر</td></tr>`:""),document.getElementById("preview-table").style.display="block"
}
function initColorPickers() {
  document.getElementById("headerBg").addEventListener("input",e=> {
    design.headerBg=e.target.value,document.getElementById("preview-header").style.background=e.target.value
  }),document.getElementById("headerText").addEventListener("input",e=> {
    design.headerText=e.target.value,document.getElementById("preview-header").style.color=e.target.value
  })
}
function pickColor(e,t) {
  const n=e.dataset.color;
  e.parentElement.querySelectorAll(".color-swatch").forEach(e=>e.classList.remove("selected")),e.classList.add("selected"),document.getElementById(t).value=n,"headerBg"===t&&(design.headerBg=n,document.getElementById("preview-header").style.background=n),"headerText"===t&&(design.headerText=n,document.getElementById("preview-header").style.color=n),autoSaveCurrentProject()
}
function setAnswerStyle() {
  const e=document.getElementById("answerStyle")?.value||"normal";
  design.answerStyle=e,autoSaveCurrentProject(),showToast(`تم تعيين نمط الإجابات: ${"highlighted"===e?"تظليل تلقائي":"جدول عادي"} ✓`,"success")
}
function updatePreview() {
  const e=e=>document.getElementById(e);
  e("prev-institution")&&(e("prev-institution").textContent=e("institution")?.value||""),e("prev-subject")&&(e("prev-subject").textContent=(e("labelSubject")?.value||"امتحان مادة:")+" "+(e("subjectName")?.value||""));
  const t=e("preview-header");
  if(t) {
    const n=t.querySelector("div:last-child");
    n&&(n.textContent=`${e("labelDuration")?.value||"الزمن:"} ${e("examDuration")?.value||""} | ${e("labelGrade")?.value||"الدرجة:"} ${e("totalGrade")?.value||""}`)
  }
  design.fontSize=parseInt(e("fontSize")?.value||14),design.layoutCols=parseInt(e("layoutCols")?.value||2),examCfg.subject=e("subjectName")?.value||"",examCfg.institution=e("institution")?.value||"",examCfg.duration=e("examDuration")?.value||"",examCfg.grade=e("totalGrade")?.value||"",examCfg.date=e("examDate")?.value||"",examCfg.instructions=e("instructions")?.value||"",examCfg.labelSubject=e("labelSubject")?.value||"",examCfg.labelDuration=e("labelDuration")?.value||"",examCfg.labelGrade=e("labelGrade")?.value||"",examCfg.labelModel=e("labelModel")?.value||"",examCfg.labelField1=e("labelField1")?.value||"",examCfg.labelField2=e("labelField2")?.value||"",examCfg.labelField3=e("labelField3")?.value||"",examCfg.labelField4=e("labelField4")?.value||"",autoSaveCurrentProject()
}
function shuffle(e) {
  const t=[...e];
  for(let e=t.length-1; e>0; e--) {
    const n=Math.floor(Math.random()*(e+1));
    [t[e],t[n]]=[t[n],t[e]]
  }
  return t
}
function deduplicateChoices(e) {
  const t=new Set;
  return e.filter(e=> {
    const n=(e.val||"").trim();
    return!!n&&!t.has(n)&&(t.add(n),!0)
  })
}
function mapEnToAr(e) {
  return {
    A:"A",B:"B",C:"C",D:"D",a:"A",b:"B",c:"C",d:"D","أ":"A","ب":"B","ج":"C","د":"D"
  }
  [e]||e.toUpperCase()||e
}
function lbl(e,t) {
  return examCfg[e]&&examCfg[e].trim()?examCfg[e].trim():t
}
function generateModels() {
  if(!questions.length&&!multiPartEnabled)return void showToast("لا توجد أسئلة للتوليد","error");
  if(multiPartEnabled&&0===mpParts.length)return void showToast("يرجى إضافة أجزاء المادة أولاً","error");
  const e=parseInt(document.getElementById("numModels").value)||4,t=parseInt(document.getElementById("qPerModel").value)||20;
  if(!multiPartEnabled&&t>questions.length)return void showToast(`عدد الأسئلة المطلوب (${t}) أكبر من المتاح (${questions.length})`,"error");
  const n=e=>document.getElementById(e)?.value||"";
  examCfg= {
    subject:n("subjectName"),institution:n("institution"),duration:n("examDuration"),grade:n("totalGrade"),classLevel:n("classLevel"),date:n("examDate"),instructions:n("instructions"),labelSubject:n("labelSubject"),labelDuration:n("labelDuration"),labelGrade:n("labelGrade"),labelModel:n("labelModel"),labelField1:n("labelField1"),labelField2:n("labelField2"),labelField3:n("labelField3"),labelField4:n("labelField4"),qCount:t,n:e
  };
  const o=multiPartEnabled?buildMpQuestionsPool():questions;
  if(!o||0===o.length)return void showToast("لا توجد أسئلة للتوليد","error");
  if(o.length<t)return void showToast(`عدد الأسئلة المتاح (${o.length}) أقل من المطلوب (${t})`,"error");
  const r=document.getElementById("gen-btn");
  r.classList.add("generating"),setTimeout(()=> {
    models=[];
    const n="أبجدهوزحطيكلمنسعفصقرشتثخذضظغ".split(""),a=new Set;
    for(let r=0; r<e; r++) {
      let e=o.map((e,t)=>( {
        q:e,idx:t
      }));
      if(opts.unique) {
        const n=e.filter(( {
          idx:e
        })=>!a.has(e));
        n.length>=t&&(e=n)
      }
      const s=shuffle(e).slice(0,t);
      s.forEach(( {
        idx:e
      })=>a.add(e));
      const l=s.map(( {
        q:e
      })=> {
        const t=e.__partColMap||colMap,n=deduplicateChoices([ {
          label:"A",val:(e[t.a]||"").trim()
        }, {
          label:"B",val:(e[t.b]||"").trim()
        }, {
          label:"C",val:(e[t.c]||"").trim()
        }, {
          label:"D",val:(e[t.d]||"").trim()
        }
        ]),o=(e[t.ans]||"").trim(),r=n.find(e=>e.val===o||e.label===o||e.label===mapEnToAr(o)),a=r?r.val:o;
        let s;
        if(opts.shuffleA&&n.length>1) {
          const e=n.map(e=>e.label),t=shuffle(n.map(e=>e.val));
          s=e.map((e,n)=>( {
            label:e,val:t[n]
          }))
        }
        else s=n;
        const l=s.find(e=>e.val===a)?.label||o;
        return {
          text:(e[t.q]||"").trim(),choices:s,correctLabel:l,topic:(e[t.topic]||e.__partName||"").trim(),diff:(e[t.diff]||"").trim()
        }
      });
      models.push( {
        name:"نموذج "+n[r],letter:n[r],questions:opts.shuffleQ?shuffle(l):l
      })
    }
    r.classList.remove("generating"),renderModelTabs(),autoSaveCurrentProject(),goPanel(3),document.getElementById("gen-summary").textContent=`${e} نماذج × ${t} سؤال = ${e*t} سؤال إجمالي`,showToast(`تم توليد ${e} نماذج بنجاح`,"success")
  },800)
}
function renderModelTabs() {
  document.getElementById("model-tabs").innerHTML=models.map((e,t)=>`<button class="exam-tab ${0===t?"active":""}" onclick="showModel(${t})">${e.name}</button>`).join("")+'<button class="exam-tab" onclick="showAnswerKeys()">نماذج الإجابات</button>',showModel(0)
}
function showModel(e) {
  currentModelIdx=e,document.querySelectorAll(".exam-tab").forEach((t,n)=>t.classList.toggle("active",n===e)),document.getElementById("answer-key-view").style.display="none",document.getElementById("model-view").style.display="block",document.getElementById("model-view").innerHTML=renderExamPaper(models[e],e)
}
function showAnswerKeys() {
  document.querySelectorAll(".exam-tab").forEach(e=>e.classList.remove("active")),document.querySelectorAll(".exam-tab").item(models.length).classList.add("active"),document.getElementById("model-view").style.display="none",document.getElementById("answer-key-view").style.display="block",renderAllAnswerKeys()
}
function renderExamPaper(e,t) {
  const n=design.fontSize,o=design.headerBg,r="highlighted"===design.answerStyle,a=opts.header?`\n    <div style="background:${o};color:${design.headerText};\n                padding:1.5rem;text-align:center;margin:-2.5rem -2.5rem 1.5rem;">\n      <div style="font-size:${n+2}px;font-weight:700;margin-bottom:4px;">${examCfg.institution}</div>\n      <div style="font-size:${n+8}px;font-weight:900;font-family:'Amiri',serif;">${lbl("labelSubject","امتحان مادة:")} ${examCfg.subject}</div>\n      <div style="font-size:${n-1}px;opacity:.85;margin-top:4px;">${lbl("labelDuration","الزمن:")} ${examCfg.duration} | ${lbl("labelGrade","الدرجة:")} ${examCfg.grade}</div>\n      ${opts.modelNum?`<div style="margin-top:6px;display:inline-block;background:rgba(255,255,255,.25);\n        padding:3px 16px;border-radius:20px;font-size:${n+1}px;font-weight:700;">$ {
    lbl("labelModel","النموذج:")
  }
  $ {
    e.letter
  }
  </div>`:""}\n    </div>`:"",s=opts.student?`\n    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:1.5rem;font-size:${n-1}px;">\n      <div style="border-bottom:1px solid #2d3748;padding-bottom:4px;color:#4a5568;font-weight:500;">${lbl("labelField1","اسم الطالب:")} ________________</div>\n      <div style="border-bottom:1px solid #2d3748;padding-bottom:4px;color:#4a5568;font-weight:500;">${lbl("labelField2","رقم الجلوس:")} ____________</div>\n      <div style="border-bottom:1px solid #2d3748;padding-bottom:4px;color:#4a5568;font-weight:500;">${lbl("labelField3","الشعبة / الفصل:")} __________</div>\n      <div style="border-bottom:1px solid #2d3748;padding-bottom:4px;color:#4a5568;font-weight:500;">التاريخ: ${examCfg.date||"___________"}</div>\n    </div>`:"",l=examCfg.instructions?`<div style="background:#f8f9ff;border:1px solid #d0d8f0;border-radius:8px;\n                   padding:10px 14px;margin-bottom:1.5rem;font-size:${n-1}px;color:#4a5568;line-height:1.8;">\n         <strong>التعليمات:</strong> ${examCfg.instructions}</div>`:"";
  let i,d="",c="";
  if(c=e.questions.map((e,t)=> {
    const o=Math.max(...e.choices.map(e=>(e.val||"").length))>50?2:4,r=[];
    for(let t=0; t<e.choices.length; t+=o) {
      const a=[];
      for(let r=t; r<Math.min(t+o,e.choices.length); r++) {
        const t=e.choices[r],s=100/o;
        a.push(`<div style="width:${s}%;display:flex;gap:6px;align-items:flex-start;font-size:${n-1}px;color:#4a5568;padding-right:8px;"><span style="font-weight:700;color:#2d3748;flex-shrink:0;min-width:18px;">${t.label})</span><span>${renderMath(t.val)}</span></div>`)
      }
      r.push(`<div style="display:flex;width:100%;margin-bottom:6px;">${a.join("")}</div>`)
    }
    const a=r.join("");
    return`\n      <div style="display:block;width:100%;margin:0 0 1.4rem 0;padding-bottom:1rem;\n                  border-bottom:1px solid #e0e4ef;page-break-inside:avoid;">\n        <div style="font-size:${n}px;line-height:1.9;margin-bottom:8px;color:#1a1a2e;font-weight:500;">\n          <span style="font-weight:700;">${t+1})</span> ${renderMath(e.text)}\n        </div>\n        <div style="padding-right:24px;">${a}</div>\n      </div>`
  }).join(""),r&&opts.answerTable) {
    const t=".omr-circle{display:inline-block;width:11px;height:11px;border:1px solid #333;border-radius:50%;margin:0 auto;background:#fff;}.omr-table{width:100%;border-collapse:collapse;border:1px solid #666;margin-bottom:1rem;page-break-inside:avoid;font-size:7pt;}.omr-table td{border:1px solid #666;padding:2px 1px;text-align:center;font-weight:600;height:16px;}.omr-table .row-label{background:#e5e5e5;font-weight:700;color:#333;min-width:22px;}.omr-table .qnum-cell{background:#f0f0f0;font-weight:700;color:#333;font-size:6.5pt;}",o=[],r=e.questions[0].choices.map(e=>e.label)||["A","B","C","D"];
    for(let t=0; t<e.questions.length; t+=20) {
      const n=Math.min(20,e.questions.length-t);
      let a='<table class="omr-table" style="width:100%;table-layout:fixed;">';
      a+='<tr><td class="row-label">Q</td>';
      for(let e=0; e<20; e++)a+=e<n?`<td class="qnum-cell">${t+e+1}</td>`:'<td class="qnum-cell" style="background:#f5f5f5;"></td>';
      a+="</tr>",r.slice(0,4).forEach(e=> {
        a+=`<tr><td class="row-label">${e}</td>`;
        for(let e=0; e<20; e++)a+=e<n?'<td style="padding:1px;"><div class="omr-circle"></div></td>':'<td style="background:#f5f5f5;"></td>';
        a+="</tr>"
      }),a+="</table>",o.push(a)
    }
    d=`<style>${t}</style><div style="padding:8px 12px;background:#f3f4f6;border:1px solid #d1d5db;border-radius:6px;margin-bottom:1rem;font-size:${n-2}px;color:#374151;line-height:1.5;">\n      <strong>ورقة الإجابات:</strong> ضع دائرة حول الإجابة الصحيحة\n    </div>${o.join("")}`,i=""
  }
  else i=opts.answerTable?buildAnswerFillTable(e.questions.length,o,n):"";
  const p="start"===opts.answerTablePos,u=detectDir(),m=`\n    <div class="exam-paper" id="paper-${t}" style="direction:${u};">\n      ${a}${s}${l}\n      ${!p||r&&opts.answerTable?"":i}\n      ${p&&r&&opts.answerTable?d:""}\n      <div style="display:block;width:100%;">${c}</div>\n      ${p?"":r&&opts.answerTable?d:i}\n    </div>`;
  return setTimeout(()=> {
    const e=document.getElementById("paper-"+t);
    e&&typesetMath(e)
  },50),m
}
function buildAnswerFillTable(e,t,n) {
  let o="";
  for(let n=0; n<e; n+=20) {
    const r=Math.min(20,e-n),a=[],s=[];
    for(let e=0; e<20; e++) {
      const t=e<r;
      a.push(t?`<td style="text-align:center;border:1px solid #c8d0e0;padding:2px 1px;min-width:28px;font-size:6.5pt;font-weight:700;color:#333;">${n+e+1}</td>`:'<td style="border:1px solid #c8d0e0;padding:2px 1px;min-width:28px;background:#f5f5f5;"></td>'),s.push(t?'<td style="border:1px solid #c8d0e0;padding:0;min-width:28px;height:16px;"></td>':'<td style="border:1px solid #c8d0e0;padding:0;min-width:28px;height:16px;background:#f5f5f5;"></td>')
    }
    o+=`\n      <tr>\n        <td style="border:1px solid #c8d0e0;padding:2px 4px;font-size:6.5pt;font-weight:700;\n                   color:#fff;background:${t};white-space:nowrap;width:48px;">Q</td>\n        ${a.join("")}\n      </tr>\n      <tr>\n        <td style="border:1px solid #c8d0e0;padding:2px 4px;font-size:6.5pt;font-weight:700;\n                   color:#fff;background:${t};white-space:nowrap;width:48px;">A</td>\n        ${s.join("")}\n      </tr>`
  }
  return`\n    <div style="margin-top:1.5rem;page-break-inside:avoid;direction:ltr;">\n      <table style="border-collapse:collapse;border:1px solid #c8d0e0;table-layout:fixed;width:100%;">${o}</table>\n    </div>`
}
function renderAllAnswerKeys() {
  const e=models.map((e,t)=> {
    const n=buildAnswerSheetHtml(e);
    return`\n      <div style="margin-bottom:2rem;border:1px solid #ddd;border-radius:8px;overflow:hidden;">\n        <div style="background:${design.headerBg};color:${design.headerText};padding:12px 16px;font-weight:700;font-size:14px;">\n          📋 ورقة الإجابات — ${e.name}\n        </div>\n        <iframe id="sheet-${t}" style="width:100%;height:600px;border:none;background:#fff;" srcdoc="${n.replace(/"/g,"&quot;")}"></iframe>\n      </div>`
  }).join("");
  document.getElementById("answer-keys-content").innerHTML=e
}
function exportPageCss(e) {
  return`\n    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&family=Amiri:wght@400;700&display=swap');\n    @page { size:A4; margin:18mm 20mm 20mm 20mm; }\n    @page { @bottom-center {\n      content: "صفحة " counter(page) " من " counter(pages);\n      font-family:'Cairo',sans-serif; font-size:9pt; color:#666;\n    }}\n    * { box-sizing:border-box; margin:0; padding:0; }\n    body { font-family:'Cairo',sans-serif; background:#fff; color:#1a1a2e; direction:rtl; font-size:13px; }\n    @media print { body { -webkit-print-color-adjust:exact; print-color-adjust:exact; } }\n    .q-block { page-break-inside: avoid; }\n    ${e||""}`
}
function buildQuestionPaperHtml(e) {
  const t=design.headerBg,n=design.headerText,o="highlighted"===design.answerStyle,r=detectDir();
  let a=e.questions.map((e,t)=> {
    const n=Math.max(...e.choices.map(e=>(e.val||"").length))>50?2:4,o=100/n,r=[];
    for(let t=0; t<e.choices.length; t+=n) {
      const a=[];
      for(let r=t; r<Math.min(t+n,e.choices.length); r++) {
        const t=e.choices[r];
        a.push(`<td style="width:${o}%;padding:2px 8px;font-size:9.5pt;color:#333;vertical-align:top;">\n          ${t?`<b>$ {
          t.label
        })</b> $ {
          renderMath(t.val)
        }
        `:""}\n        </td>`)
      }
      r.push(`<tr>${a.join("")}</tr>`)
    }
    return`<div class="q-block" style="margin-bottom:9px;padding-bottom:8px;border-bottom:1px solid #dde2ee;">\n      <div style="font-size:10.5pt;line-height:1.75;color:#1a1a2e;margin-bottom:4px;">\n        <b>${t+1})</b> ${renderMath(e.text)}\n      </div>\n      <table style="width:100%;border-collapse:collapse;">${r.join("")}</table>\n    </div>`
  }).join(""),s="",l="";
  if(o&&opts.answerTable) {
    const t=".omr-circle{display:inline-block;width:11px;height:11px;border:1px solid #333;border-radius:50%;margin:0 auto;background:#fff;}.omr-table{width:100%;border-collapse:collapse;border:1px solid #666;margin-bottom:10pt;page-break-inside:avoid;font-size:7pt;}.omr-table td{border:1px solid #666;padding:2px 1px;text-align:center;font-weight:600;height:15px;}.omr-table .row-label{background:#e5e5e5;font-weight:700;color:#333;min-width:20px;}.omr-table .qnum-cell{background:#f0f0f0;font-weight:700;color:#333;font-size:6.5pt;}",n=[],o=e.questions[0].choices.map(e=>e.label)||["A","B","C","D"];
    for(let t=0; t<e.questions.length; t+=20) {
      const r=Math.min(20,e.questions.length-t);
      let a='<table class="omr-table" style="width:100%;table-layout:fixed;">';
      a+='<tr><td class="row-label">Q</td>';
      for(let e=0; e<20; e++)a+=e<r?`<td class="qnum-cell">${t+e+1}</td>`:'<td class="qnum-cell" style="background:#f5f5f5;"></td>';
      a+="</tr>",o.slice(0,4).forEach(e=> {
        a+=`<tr><td class="row-label">${e}</td>`;
        for(let e=0; e<20; e++)a+=e<r?'<td style="padding:1px;"><div class="omr-circle"></div></td>':'<td style="background:#f5f5f5;"></td>';
        a+="</tr>"
      }),a+="</table>",n.push(a)
    }
    s=`<style>${t}</style><div style="padding:7px 10px;background:#f3f4f6;border:1px solid #d1d5db;border-radius:4px;margin-bottom:10pt;font-size:8pt;color:#374151;line-height:1.4;">\n      <strong>ورقة الإجابات:</strong> ضع دائرة حول الإجابة الصحيحة\n    </div>${n.join("")}`
  }
  else if(opts.answerTable) {
    const n=20;
    let o="";
    for(let r=0; r<e.questions.length; r+=n) {
      const a=Math.min(n,e.questions.length-r),s=[],l=[];
      for(let e=0; e<n; e++) {
        const t=e<a;
        s.push(t?`<td style="text-align:center;border:1px solid #c8d0e0;padding:2px 1px;min-width:28px;font-size:6.5pt;font-weight:700;color:#333;">${r+e+1}</td>`:'<td style="border:1px solid #c8d0e0;padding:2px 1px;min-width:28px;background:#f5f5f5;"></td>'),l.push(t?'<td style="border:1px solid #c8d0e0;min-width:28px;height:16px;"></td>':'<td style="border:1px solid #c8d0e0;min-width:28px;height:16px;background:#f5f5f5;"></td>')
      }
      o+=`<tr>\n          <td style="border:1px solid #c8d0e0;padding:2px 4px;font-size:6.5pt;font-weight:700;\n                     color:#fff;background:${t};white-space:nowrap;width:48px;">Q</td>${s.join("")}\n        </tr><tr>\n          <td style="border:1px solid #c8d0e0;padding:2px 4px;font-size:6.5pt;font-weight:700;\n                     color:#fff;background:${t};white-space:nowrap;width:48px;">A</td>${l.join("")}\n        </tr>`
    }
    l=`<div style="margin-top:14pt;page-break-inside:avoid;direction:ltr;">\n        <table style="border-collapse:collapse;border:1px solid #c8d0e0;table-layout:fixed;width:100%;">${o}</table>\n      </div>`
  }
  const i="start"===opts.answerTablePos,d=i&&o&&opts.answerTable?s+a:a,c=!i&&o&&opts.answerTable?s:l;
  return`<!DOCTYPE html><html lang="${"ltr"===r?"en":"ar"}" dir="${r}"><head><meta charset="UTF-8">\n<style>${exportPageCss()}</style>\n${mathJaxScript()}\n</head><body style="direction:${r};">\n  <div style="background:${t};color:${n};padding:14px 20px;text-align:center;margin-bottom:12px;">\n    <div style="font-size:13pt;font-weight:700;">${examCfg.institution||""}</div>\n    <div style="font-size:17pt;font-weight:900;font-family:'Amiri',serif;margin:4px 0;">${lbl("labelSubject","امتحان مادة:")} ${examCfg.subject||""}</div>\n    <div style="font-size:10pt;opacity:.9;">${lbl("labelDuration","الزمن:")} ${examCfg.duration} | ${lbl("labelGrade","الدرجة:")} ${examCfg.grade} | ${examCfg.date||""}</div>\n    <div style="margin-top:5px;display:inline-block;background:rgba(255,255,255,.25);\n                padding:2px 18px;border-radius:20px;font-size:11pt;font-weight:700;">${lbl("labelModel","النموذج:")} ${e.letter}</div>\n  </div>\n  <table style="width:100%;border-collapse:collapse;font-size:10pt;margin-bottom:10px;"><tr>\n    <td style="padding:4px 8px;border-bottom:1px solid #333;width:50%;">${lbl("labelField1","اسم الطالب:")} _______________________</td>\n    <td style="padding:4px 8px;border-bottom:1px solid #333;">${lbl("labelField2","رقم الجلوس:")} ________________</td>\n  </tr><tr>\n    <td style="padding:4px 8px;border-bottom:1px solid #333;">${lbl("labelField3","الشعبة / الفصل:")} __________________</td>\n    <td style="padding:4px 8px;border-bottom:1px solid #333;">${lbl("labelField4","التاريخ:")} ${examCfg.date||"_______________"}</td>\n  </tr></table>\n  ${examCfg.instructions?`<div style="background:#f5f7ff;border:1px solid #c8d0f0;border-radius:5px;\n    padding:7px 12px;margin-bottom:10px;font-size:9.5pt;color:#333;line-height:1.7;">\n    <strong>التعليمات:</strong> $ {
    examCfg.instructions
  }
  </div>`:""}\n  ${!i||o&&opts.answerTable?"":l}\n  ${d}\n  ${i?"":c}\n</body></html>`
}
function buildAnswerSheetHtml(e) {
  const t=design.headerBg,n=design.headerText;
  if("highlighted"!==design.answerStyle) {
    const o=e.questions.map((e,t)=> {
      const n=Math.max(...e.choices.map(e=>(e.val||"").length))>50?2:4,o=100/n,r=e.choices.map(e=>`<td style="width:${o}%;padding:2px 8px;font-size:9.5pt;vertical-align:top;color:#333;">\n          <b style="flex-shrink:0;">${e.label})</b> ${renderMath(e.val)}</td>`),a=[];
      for(let e=0; e<r.length; e+=n) {
        const t=[];
        for(let a=e; a<Math.min(e+n,r.length); a++)t.push(r[a]||`<td style="width:${o}%;"></td>`);
        a.push(`<tr>${t.join("")}</tr>`)
      }
      return`<div class="q-block" style="margin-bottom:9px;padding-bottom:8px;border-bottom:1px solid #dde2ee;">\n        <div style="font-size:10.5pt;line-height:1.75;color:#1a1a2e;margin-bottom:4px;">\n          <b>${t+1})</b> ${renderMath(e.text)}\n        </div>\n        <table style="width:100%;border-collapse:collapse;margin-right:16px;">${a.join("")}</table>\n      </div>`
    }).join(""),r=e.questions.map((e,t)=>`<td style="text-align:center;border:1px solid #d1d5db;padding:5px 3px;min-width:40px;background:#f3f4f6;">\n        <div style="font-size:8pt;color:#666;">${t+1}</div>\n        <div style="font-size:12pt;font-weight:700;color:#374151;">${e.correctLabel}</div>\n      </td>`),a=[];
    for(let e=0; e<r.length; e+=20)a.push(`<tr>${r.slice(e,e+20).join("")}</tr>`);
    const s=detectDir();
    return`<!DOCTYPE html><html lang="${"ltr"===s?"en":"ar"}" dir="${s}"><head><meta charset="UTF-8">\n<style>${exportPageCss()}</style>\n${mathJaxScript()}\n</head><body style="direction:${s};">\n  <div style="background:${t};color:${n};padding:14px 20px;text-align:center;margin-bottom:12px;">\n    <div style="font-size:13pt;font-weight:700;">${examCfg.institution||""}</div>\n    <div style="font-size:17pt;font-weight:900;font-family:'Amiri',serif;margin:4px 0;">${lbl("labelSubject","امتحان مادة:")} ${examCfg.subject||""}</div>\n    <div style="font-size:10pt;opacity:.9;">${lbl("labelDuration","الزمن:")} ${examCfg.duration} | ${lbl("labelGrade","الدرجة:")} ${examCfg.grade} | ${examCfg.date||""}</div>\n    <div style="margin-top:5px;display:inline-block;background:rgba(255,255,255,.25);padding:2px 18px;border-radius:20px;font-size:11pt;font-weight:700;">${lbl("labelModel","النموذج:")} ${e.letter}</div>\n  </div>\n  ${o}\n  <div style="margin-top:18pt;page-break-inside:avoid;">\n    <div style="background:#4b5563;color:#fff;padding:5pt 12pt;font-size:10pt;font-weight:700;border-radius:4pt 4pt 0 0;display:inline-block;">ملخص الإجابات الصحيحة</div>\n    <table style="width:100%;border-collapse:collapse;border:1px solid #d1d5db;">${a.join("")}</table>\n  </div>\n</body></html>`
  }
  const o=[],r=e.questions[0]?e.questions[0].choices.map(e=>e.label):["A","B","C","D"];
  for(let t=0; t<e.questions.length; t+=20) {
    const n=Math.min(20,e.questions.length-t);
    let a='<table class="omr-table" style="width:100%;table-layout:fixed;">';
    a+='<tr><td class="row-label">السؤال</td>';
    for(let e=0; e<20; e++)a+=e<n?`<td class="qnum-cell">${t+e+1}</td>`:'<td class="qnum-cell" style="background:#f5f5f5;"></td>';
    a+="</tr>",r.forEach((o,r)=> {
      a+=`<tr><td class="row-label">${o}</td>`;
      for(let o=0; o<20; o++)if(o<n) {
        const n=e.questions[t+o],s=(n&&n.choices||[])[r],l=s&&s.label===n.correctLabel;
        a+=`<td style="padding:8px 4px;"><div class="${l?"omr-circle filled":"omr-circle"}"></div></td>`
      }
      else a+='<td style="padding:8px 4px;background:#f5f5f5;"></td>';
      a+="</tr>"
    }),a+="</table>",o.push(a)
  }
  const a=e.questions.map((e,t)=>`\n    <td style="text-align:center;border:1px solid #d1d5db;padding:8px 4px;min-width:40px;background:#f9fafb;vertical-align:middle;">\n      <div style="font-size:8pt;color:#666;margin-bottom:2px;">${t+1}</div>\n      <div style="font-size:12pt;font-weight:700;color:#1f2937;">${e.correctLabel}</div>\n    </td>`),s=[];
  for(let e=0; e<a.length; e+=20)s.push(`<tr>${a.slice(e,e+20).join("")}</tr>`);
  const l=detectDir();
  return`<!DOCTYPE html><html lang="${"ltr"===l?"en":"ar"}" dir="${l}"><head><meta charset="UTF-8">\n<style>${exportPageCss("\n    .omr-circle {\n      display: inline-block;\n      width: 14px;\n      height: 14px;\n      border: 1.5px solid #333;\n      border-radius: 50%;\n      margin: 0 auto;\n      background: #fff;\n    }\n    .omr-circle.filled {\n      background: #333;\n    }\n    .omr-table {\n      width: 100%;\n      border-collapse: collapse;\n      border: 1px solid #333;\n      margin-bottom: 14pt;\n      page-break-inside: avoid;\n    }\n    .omr-table td {\n      border: 1px solid #333;\n      padding: 6px;\n      text-align: center;\n      font-size: 9pt;\n      font-weight: 600;\n    }\n    .omr-table .row-label {\n      background: #f3f4f6;\n      font-weight: 700;\n      color: #1f2937;\n      min-width: 30px;\n    }\n    .omr-table .qnum-cell {\n      background: #f9fafb;\n      font-weight: 700;\n      color: #374151;\n    }\n  ")}</style>\n${mathJaxScript()}\n</head><body style="direction:${l};">\n  <div style="background:${t};color:${n};padding:14px 20px;text-align:center;margin-bottom:12px;">\n    <div style="font-size:13pt;font-weight:700;">${examCfg.institution||""}</div>\n    <div style="font-size:17pt;font-weight:900;font-family:'Amiri',serif;margin:4px 0;">${lbl("labelSubject","امتحان مادة:")} ${examCfg.subject||""}</div>\n    <div style="font-size:10pt;opacity:.9;">${lbl("labelDuration","الزمن:")} ${examCfg.duration} | ${lbl("labelGrade","الدرجة:")} ${examCfg.grade} | ${examCfg.date||""}</div>\n    <div style="margin-top:5px;display:inline-flex;gap:8px;align-items:center;justify-content:center;flex-wrap:wrap;">\n      <span style="display:inline-block;background:rgba(255,255,255,.25);padding:2px 18px;border-radius:20px;font-size:11pt;font-weight:700;">${lbl("labelModel","النموذج:")} ${e.letter}</span>\n      <span style="display:inline-block;background:#1f2937;color:#fff;padding:2px 18px;border-radius:20px;font-size:10pt;font-weight:700;">جدول OMR</span>\n    </div>\n  </div>\n  \n  <div style="padding:10px 15px;background:#f3f4f6;border:1px solid #d1d5db;border-radius:6px;margin-bottom:12px;font-size:9pt;color:#374151;line-height:1.6;">\n    <strong>التعليمات:</strong> املأ الدائرة المطابقة لكل إجابة صحيحة. الدوائر المملوءة تمثل الإجابات الصحيحة.\n  </div>\n  \n  ${o.join("")}\n  \n  <div style="margin-top:18pt;page-break-inside:avoid;">\n    <div style="background:#1f2937;color:#fff;padding:8pt 12pt;font-size:10pt;font-weight:700;border-radius:4pt 4pt 0 0;display:inline-block;">✓ مفتاح الإجابات الصحيحة</div>\n    <table style="width:100%;border-collapse:collapse;border:1px solid #d1d5db;">${s.join("")}</table>\n  </div>\n</body></html>`
}
async function exportAllPDF() {
  if(!models.length)return void showToast("لا توجد نماذج بعد","error");
  showToast("جاري إنشاء ملفات PDF...","");
  const e=new JSZip,t=e.folder("الأسئلة"),n=e.folder("نماذج_الإجابات");
  models.forEach(e=> {
    t.file(`${e.name} - أسئلة.html`,buildQuestionPaperHtml(e)),n.file(`${e.name} - إجابات.html`,buildAnswerSheetHtml(e))
  });
  let o=`<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="UTF-8">\n<style>\n  ${exportPageCss()}\n  .model-wrap { page-break-after: always; }\n</style></head><body>`;
  models.forEach(e=> {
    const t=buildQuestionPaperHtml(e).replace(/[\s\S]*?<body[^>]*>/i,"").replace(/<\/body>[\s\S]*/i,"");
    o+=`<div class="model-wrap">${t}</div>`
  }),o+="</body></html>",e.file("كل_النماذج_مجمعة.html",o),triggerDownload(await e.generateAsync( {
    type:"blob"
  }),`نماذج_PDF_${examCfg.subject||"امتحان"}.zip`),showToast("ZIP جاهز — افتح HTML واضغط Ctrl+P للطباعة كـ PDF","success")
}
async function exportAllWord() {
  if(!models.length)return void showToast("لا توجد نماذج بعد","error");
  showToast("جاري إنشاء ملفات Word...","");
  const e=new JSZip,t=e.folder("الأسئلة"),n=e.folder("نماذج_الإجابات"),o=design.headerBg,r=design.headerText,a=detectDir(),s=(e,t,n)=>`<html xmlns:o="urn:schemas-microsoft-com:office:office"\n           xmlns:w="urn:schemas-microsoft-com:office:word"\n           xmlns="http://www.w3.org/TR/REC-html40">\n<head><meta charset="utf-8"><style>\n  @page Section1 { size:21cm 29.7cm; margin:18mm 20mm 20mm 20mm;\n    mso-header-margin:10mm; mso-footer-margin:10mm; mso-page-numbers:1; }\n  div.Section1 { page:Section1; }\n  body { font-family:'Arial Unicode MS',Arial,sans-serif; direction:${a}; font-size:10pt; }\n  p { margin:0 0 3pt; text-align:${"ltr"===a?"left":"right"}; }\n  td { text-align:${"ltr"===a?"left":"right"}; }\n  .correct { background:#bbf7d0; border:1pt solid #16a34a; padding:1pt 4pt; }\n</style></head>\n<body dir="${a}"><div class="Section1">\n  <div style="background:${o};color:${r};padding:10pt;text-align:center;margin-bottom:10pt;\n              ${n?"border:3pt solid #16a34a;":""}">\n    <div style="font-size:13pt;font-weight:700;">${examCfg.institution||""}</div>\n    <div style="font-size:17pt;font-weight:900;">${lbl("labelSubject","امتحان مادة:")} ${examCfg.subject||""}</div>\n    <div style="font-size:10pt;">${lbl("labelDuration","الزمن:")} ${examCfg.duration} &nbsp;|&nbsp; ${lbl("labelGrade","الدرجة:")} ${examCfg.grade} &nbsp;|&nbsp; ${examCfg.date||""}</div>\n    <div style="font-size:11pt;font-weight:700;">${lbl("labelModel","النموذج:")} ${t}${n?" — Answer Key":""}</div>\n  </div>\n  ${e}\n</div></body></html>`;
  models.forEach(e=> {
    const r=`<table width="100%" style="margin-bottom:8pt;" dir="${a}"><tr>\n      <td width="50%" style="border-bottom:1pt solid #333;padding:3pt;font-size:10pt;">${lbl("labelField1","اسم الطالب:")} _______________________</td>\n      <td style="border-bottom:1pt solid #333;padding:3pt;font-size:10pt;">${lbl("labelField2","رقم الجلوس:")} ________________</td>\n    </tr><tr>\n      <td style="border-bottom:1pt solid #333;padding:3pt;font-size:10pt;">${lbl("labelField3","الشعبة / الفصل:")} __________________</td>\n      <td style="border-bottom:1pt solid #333;padding:3pt;font-size:10pt;">${lbl("labelField4","التاريخ:")} ${examCfg.date||"___"}</td>\n    </tr></table>`,l=examCfg.instructions?`<p style="background:#f5f7ff;border:1pt solid #c0c8e0;padding:5pt;font-size:9.5pt;margin-bottom:8pt;">\n           <b>التعليمات:</b> ${examCfg.instructions}</p>`:"",i=e.questions.map((e,t)=> {
      const n=Math.max(...e.choices.map(e=>(e.val||"").length))>50?2:4,o=100/n,r=[];
      for(let t=0; t<e.choices.length; t+=n) {
        const a=[];
        for(let r=t; r<Math.min(t+n,e.choices.length); r++) {
          const t=e.choices[r];
          a.push(`<td width="${o}%" style="padding:2pt 6pt;font-size:10pt;">${t?`<b>$ {
            t.label
          })</b> $ {
            t.val
          }
          `:""}</td>`)
        }
        r.push(`<tr>${a.join("")}</tr>`)
      }
      return`<div dir="${a}" style="margin-bottom:6pt;">\n        <p style="margin:0 0 4pt;font-size:10.5pt;line-height:1.75;"><b>${t+1})</b> ${e.text}</p>\n        <table width="100%" dir="${a}" style="border-collapse:collapse;margin-bottom:6pt;">${r.join("")}</table>\n        <hr style="border:none;border-top:1px solid #dde;margin:0 0 4pt;">\n      </div>`
    }).join("");
    let d="";
    if(opts.answerTable) {
      const t=20,n="start"===opts.answerTablePos?"start":"end";
      if("highlighted"===design.answerStyle) {
        const n=e.questions[0]?.choices?.map(e=>e.label)||["A","B","C","D"],o=10,r=((170-o)/t).toFixed(1);
        let a="";
        for(let s=0; s<e.questions.length; s+=t) {
          const l=Math.min(t,e.questions.length-s);
          let i='<table dir="ltr" style="border-collapse:collapse;border:1pt solid #555;margin-bottom:6pt;direction:ltr;width:100%;table-layout:fixed;">';
          i+=`<tr><td style="border:1pt solid #555;padding:3pt 4pt;font-size:7.5pt;font-weight:700;background:#d0d4de;text-align:center;width:${o}mm;">Q</td>`;
          for(let e=0; e<t; e++) {
            const t=e<l;
            i+=`<td style="border:1pt solid #555;padding:3pt 1pt;font-size:7pt;font-weight:700;text-align:center;background:${t?"#e8eaf0":"#f5f5f5"};width:${r}mm;">${t?s+e+1:""}</td>`
          }
          i+="</tr>",n.slice(0,4).forEach(e=> {
            i+=`<tr><td style="border:1pt solid #555;padding:4pt 4pt;font-size:7.5pt;font-weight:700;background:#d0d4de;text-align:center;">${e}</td>`;
            for(let e=0; e<t; e++)i+=e<l?'<td style="border:1pt solid #555;padding:4pt 1pt;text-align:center;font-size:11pt;line-height:1;height:16pt;">&#9711;</td>':'<td style="border:1pt solid #555;background:#f5f5f5;"></td>';
            i+="</tr>"
          }),i+="</table>",a+=i
        }
        d=`<br><p style="font-size:8.5pt;color:#374151;margin-bottom:5pt;"><strong>ورقة الإجابات:</strong> ضع دائرة حول الإجابة الصحيحة</p>${a}`
      }
      else {
        let n="";
        for(let r=0; r<e.questions.length; r+=t) {
          const a=Math.min(t,e.questions.length-r),s=[],l=[];
          for(let e=0; e<t; e++) {
            const t=e<a;
            s.push(t?`<td style="text-align:center;border:1pt solid #c8d0e0;padding:2pt;width:20pt;font-size:7pt;font-weight:700;">${r+e+1}</td>`:'<td style="border:1pt solid #c8d0e0;width:20pt;background:#f5f5f5;"></td>'),l.push(t?'<td style="border:1pt solid #c8d0e0;width:20pt;height:12pt;"></td>':'<td style="border:1pt solid #c8d0e0;width:20pt;height:12pt;background:#f5f5f5;"></td>')
          }
          n+=`<tr>\n            <td style="border:1pt solid #c8d0e0;padding:2pt 4pt;font-size:7pt;font-weight:700;\n                       background:${o};color:#fff;white-space:nowrap;width:30pt;">Q</td>${s.join("")}\n          </tr><tr>\n            <td style="border:1pt solid #c8d0e0;padding:2pt 4pt;font-size:7pt;font-weight:700;\n                       background:${o};color:#fff;white-space:nowrap;width:30pt;">A</td>${l.join("")}\n          </tr>`
        }
        d=`<br><table dir="ltr" style="border-collapse:collapse;direction:ltr;table-layout:fixed;width:100%;">${n}</table>`
      }
      "start"===n&&(d+="\x3c!--FILLTABLE_START--\x3e")
    }
    const c=d.includes("\x3c!--FILLTABLE_START--\x3e")?d.replace("\x3c!--FILLTABLE_START--\x3e","")+r+l+i:r+l+i+d;
    t.file(`${e.name} - أسئلة.doc`,"\ufeff"+s(c,e.letter,!1));
    const p=e.questions.map((e,t)=> {
      const n=Math.max(...e.choices.map(e=>(e.val||"").length))>50?2:4,o=100/n,r=[];
      for(let t=0; t<e.choices.length; t+=n) {
        const a=[];
        for(let r=t; r<Math.min(t+n,e.choices.length); r++) {
          const t=e.choices[r];
          if(!t)continue;
          const n=t.label===e.correctLabel;
          a.push(`<td width="${o}%" style="padding:2pt 6pt;font-size:10pt;">\n            <span${n?' class="correct" style="background:#bbf7d0;padding:1pt 4pt;"':""}>\n              <b${n?' style="color:#15803d;"':""}>${t.label})</b> ${t.val}\n            </span></td>`)
        }
        r.push(`<tr>${a.join("")}</tr>`)
      }
      return`<div dir="${a}" style="margin-bottom:6pt;">\n        <p style="margin:0 0 4pt;font-size:10.5pt;line-height:1.75;"><b>${t+1})</b> ${e.text}</p>\n        <table width="100%" dir="${a}" style="border-collapse:collapse;margin-bottom:6pt;">${r.join("")}</table>\n        <hr style="border:none;border-top:1px solid #dde;margin:0 0 4pt;">\n      </div>`
    }).join(""),u=e.questions.map((e,t)=>`<td style="text-align:center;border:1pt solid #bbf7d0;padding:4pt;width:40pt;background:#f0fdf4;">\n        <div style="font-size:8pt;color:#666;">${t+1}</div>\n        <div style="font-size:12pt;font-weight:700;color:#15803d;">${e.correctLabel}</div>\n      </td>`),m=[];
    for(let e=0; e<u.length; e+=10)m.push(`<tr>${u.slice(e,e+10).join("")}</tr>`);
    const g=`<br><div style="background:#16a34a;color:#fff;padding:4pt 10pt;font-size:10pt;font-weight:700;">ملخص الإجابات</div>\n      <table style="border-collapse:collapse;width:100%;">${m.join("")}</table>`;
    n.file(`${e.name} - إجابات.doc`,"\ufeff"+s(p+g,e.letter,!0))
  }),triggerDownload(await e.generateAsync( {
    type:"blob"
  }),`نماذج_Word_${examCfg.subject||"امتحان"}.zip`),showToast("ZIP جاهز — كل نموذج في ملفين (أسئلة + إجابات)","success")
}
async function exportAnswerKeys() {
  if(!models.length)return void showToast("لا توجد نماذج بعد","error");
  const e=new JSZip,t=e.folder("نماذج_الإجابات");
  models.forEach(e=>t.file(`${e.name} - إجابات.html`,buildAnswerSheetHtml(e))),triggerDownload(await e.generateAsync( {
    type:"blob"
  }),`نماذج_الإجابات_${examCfg.subject||"امتحان"}.zip`),showToast("تم تنزيل نماذج الإجابات","success")
}
function triggerDownload(e,t) {
  const n=document.createElement("a");
  n.href=URL.createObjectURL(e),n.download=t,n.click(),setTimeout(()=>URL.revokeObjectURL(n.href),5e3)
}
function showToast(e,t="") {
  const n=document.getElementById("toast");
  n.textContent=e,n.className="toast show "+t,setTimeout(()=>n.classList.remove("show"),3500)
}
function showProjectManager() {
  const e=getAllProjects();
  let t="";
  t=0===e.length?'\n      <div style="text-align:center;padding:2rem;color:var(--text3);">\n        <div style="font-size:48px;margin-bottom:1rem;">📭</div>\n        <p>لا توجد مشاريع محفوظة</p>\n      </div>':`\n      <div style="display:grid;gap:10px;">\n        ${e.map(e=>`\n          <div style="display:flex;align-items:center;justify-content:space-between;\n                     padding:12px;background:var(--bg3);border-radius:8px;border:1px solid var(--border);">\n            <div style="flex:1;min-width:0;">\n              <div style="font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">\n                $ {
    e.name
  }
  \n              </div>\n              <div style="font-size:12px;color:var(--text3);margin-top:4px;">\n                $ {
    e.questions
  }
  سؤال • $ {
    new Date(e.date).toLocaleString("ar-EG")
  }
  \n              </div>\n            </div>\n            <div style="display:flex;gap:6px;margin-right:1rem;">\n              <button onclick="loadProject(${e.id})" style="padding:6px 12px;background:var(--accent);color:#fff;\n                     border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600;">فتح</button>\n              <button onclick="deleteProjectConfirm(${e.id})" style="padding:6px 12px;background:#ff4f6a44;\n                     color:#ff4f6a;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600;">حذف</button>\n            </div>\n          </div>\n        `).join("")}\n      </div>`;
  const n=`\n    <div style="position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.5);\n                display:flex;align-items:center;justify-content:center;z-index:9999;" id="manager-overlay" onclick="closeProjectManager()">\n      <div style="background:var(--bg2);border-radius:12px;padding:2rem;max-width:600px;width:90%;max-height:80vh;\n                  overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.3);" onclick="event.stopPropagation()">\n        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:1.5rem;">\n          <h2 style="font-size:20px;font-weight:700;color:var(--text);margin:0;">📁 إدارة المشاريع</h2>\n          <button onclick="closeProjectManager()" style="background:transparent;border:none;font-size:24px;cursor:pointer;color:var(--text2);">✕</button>\n        </div>\n        <div style="background:var(--bg3);border:1px solid var(--border);border-radius:8px;padding:10px 12px;\n                    margin-bottom:1rem;font-size:12px;color:var(--text2);">\n          💾 <strong>موقع الحفظ:</strong> localStorage في المتصفح (آمن ومحلي)<br>\n          📍 <strong>المسار:</strong> AppData/Local/[Browser]/User Data/Local Storage\n        </div>\n        ${t}\n      </div>\n    </div>`;
  document.body.insertAdjacentHTML("beforeend",n)
}
function closeProjectManager() {
  const e=document.getElementById("manager-overlay");
  e&&e.remove()
}
function loadProject(e) {
  const t=loadProjectFromStorage(e);
  if(t) {
    questions=t.questions||[],colMap=t.colMap||colMap,opts=t.opts||opts,design=t.design||design,examCfg=t.examCfg|| {

    },models=t.models||[],window.currentProjectId=e;
    const n=(e,t)=> {
      const n=document.getElementById(e);
      n&&(n.value=t||"")
    };
    n("subjectName",examCfg.subject),n("institution",examCfg.institution),n("examDuration",examCfg.duration),n("totalGrade",examCfg.grade),n("classLevel",examCfg.classLevel),n("examDate",examCfg.date),n("instructions",examCfg.instructions),updateSidebar(),questions.length>0&&renderPreviewTable(Object.keys(colMap).filter(e=>colMap[e])),models.length>0?(goPanel(3),setTimeout(()=> {
      renderModelTabs(),showToast(`تم تحميل "${examCfg.subject||"المشروع"}" — ${models.length} نماذج ✓`,"success")
    },100)):(goPanel(0),showToast(`تم تحميل "${examCfg.subject||"المشروع"}" بنجاح ✓`,"success")),closeProjectManager()
  }
  else showToast("خطأ في تحميل المشروع","error")
}
function deleteProjectConfirm(e) {
  confirm("هل أنت متأكد من حذف هذا المشروع؟")&&(deleteProjectFromStorage(e),showToast("تم حذف المشروع ✓","success"),closeProjectManager(),setTimeout(showProjectManager,300))
}
let currentTheme="dark";
function toggleTheme() {
  currentTheme="dark"===currentTheme?"light":"dark",applyTheme(currentTheme);
  const e=document.getElementById("theme-btn");
  e&&(e.textContent="dark"===currentTheme?"☀️ Light":"🌙 Dark")
}
function applyTheme(e) {
  const t=document.documentElement;
  "light"===e?(t.style.setProperty("--bg","#f0f2f8"),t.style.setProperty("--bg2","#e4e8f2"),t.style.setProperty("--bg3","#d8ddf0"),t.style.setProperty("--card","#ffffff"),t.style.setProperty("--border","#c8d0e8"),t.style.setProperty("--border2","#a0aace"),t.style.setProperty("--text","#1a1f35"),t.style.setProperty("--text2","#4a5580"),t.style.setProperty("--text3","#7a85aa")):(t.style.setProperty("--bg","#0f1117"),t.style.setProperty("--bg2","#161b27"),t.style.setProperty("--bg3","#1e2535"),t.style.setProperty("--card","#1a2030"),t.style.setProperty("--border","#2a3347"),t.style.setProperty("--border2","#3a4a67"),t.style.setProperty("--text","#e8edf8"),t.style.setProperty("--text2","#8a9bc0"),t.style.setProperty("--text3","#5a6a8a"))
}
const RESULTS_STORAGE_KEY="omr_results_v1";
let scannerState= {
  selectedSubject:null,selectedModelIdx:null,selectedModel:null,currentStudentId:null,currentScannedAnswers:null,pendingManualQuestions:[],cameraStream:null,facingMode:"environment",allResults:[]
};
function loadScanResults() {
  try {
    const e=localStorage.getItem("omr_results_v1");
    return e?JSON.parse(e):[]
  }
  catch(e) {
    return[]
  }
}
function saveScanResults(e) {
  try {
    localStorage.setItem("omr_results_v1",JSON.stringify(e))
  }
  catch(e) {
    console.error("Save results error:",e)
  }
}
function initScannerPanel() {
  const e=getAllProjects(),t=document.getElementById("scanner-subject-select");
  if(t) {
    if(t.innerHTML='<option value="">-- اختر المادة --</option>',models.length>0&&examCfg.subject) {
      const e=document.createElement("option");
      e.value="__current__",e.textContent=`${examCfg.subject} (الجلسة الحالية — ${models.length} نماذج)`,t.appendChild(e)
    }
    e.forEach(e=> {
      const n=document.createElement("option");
      n.value=e.id,n.textContent=`${e.name} (${e.models} نماذج — ${new Date(e.date).toLocaleDateString("ar-EG")})`,t.appendChild(n)
    }),scannerState.allResults=loadScanResults(),scannerState.allResults.length>0&&renderResultsTable()
  }
}
function onScannerSubjectChange() {
  const e=document.getElementById("scanner-subject-select").value,t=document.getElementById("scanner-model-select"),n=document.getElementById("scanner-student-card"),o=document.getElementById("scanner-model-info");
  if(t.innerHTML='<option value="">-- اختر النموذج --</option>',n.style.display="none",o.style.display="none",scannerState.selectedSubject=null,!e)return;
  let r=[],a="";
  if("__current__"===e)r=models,a=examCfg.subject;
  else {
    const t=loadProjectFromStorage(parseInt(e));
    t&&(r=t.models||[],a=t.examCfg?.subject||"غير محدد")
  }
  if(0===r.length)return void showToast("لا توجد نماذج لهذه المادة","error");
  scannerState.selectedSubject= {
    val:e,models:r,name:a
  },r.forEach((e,n)=> {
    const o=document.createElement("option");
    o.value=n,o.textContent=`${e.name} — ${e.questions.length} سؤال`,t.appendChild(o)
  });
  const s=document.createElement("option");
  s.value="auto",s.textContent="🔍 تحديد تلقائي (من الورقة)",t.insertBefore(s,t.children[1]),n.style.display="block",document.getElementById("results-subject-label").textContent=a
}
function onScannerModelChange() {
  const e=document.getElementById("scanner-model-select").value,t=document.getElementById("scanner-model-info");
  if(!e||!scannerState.selectedSubject)return void(t.style.display="none");
  if("auto"===e)return t.style.display="block",t.innerHTML="🔍 سيتم تحديد النموذج تلقائياً من الورقة. إذا لم يتمكن الماسح من تحديده، سيطلب منك الاختيار يدوياً.",void(scannerState.selectedModelIdx="auto");
  const n=parseInt(e),o=scannerState.selectedSubject.models[n];
  scannerState.selectedModelIdx=n,scannerState.selectedModel=o,t.style.display="block",t.innerHTML=`✅ النموذج: <strong style="color:var(--accent);">${o.name}</strong> | عدد الأسئلة: <strong>${o.questions.length}</strong> | مفتاح الإجابات جاهز للمقارنة`
}
function startScanSession() {
  const e=document.getElementById("scanner-student-id").value.trim();
  if(e)if(document.getElementById("scanner-model-select").value) {
    if(scannerState.allResults.find(t=>t.studentId===e&&t.subject===scannerState.selectedSubject?.name)) {
      if(!confirm(`الرقم الجامعي "${e}" موجود مسبقاً لهذه المادة. هل تريد الاستبدال؟`))return;
      scannerState.allResults=scannerState.allResults.filter(t=>!(t.studentId===e&&t.subject===scannerState.selectedSubject?.name))
    }
    scannerState.currentStudentId=e,document.getElementById("scanner-current-student").textContent=e,document.getElementById("scanner-camera-card").style.display="block",document.getElementById("scan-result-content").textContent="اضغط على زر المسح لبدء العملية...",document.getElementById("manual-override-panel").style.display="none",document.getElementById("score-summary-panel").style.display="none",startCamera()
  }
  else showToast("يرجى اختيار النموذج","error");
  else showToast("يرجى إدخال الرقم الجامعي","error")
}
async function startCamera() {
  try {
    scannerState.cameraStream&&scannerState.cameraStream.getTracks().forEach(e=>e.stop());
    const e=await navigator.mediaDevices.getUserMedia( {
      video: {
        facingMode:scannerState.facingMode,width: {
          ideal:1280
        },height: {
          ideal:960
        }
      }
    });
    scannerState.cameraStream=e,document.getElementById("scanner-video").srcObject=e
  }
  catch(e) {
    showToast("لا يمكن الوصول للكاميرا: "+e.message,"error")
  }
}
function stopCamera() {
  scannerState.cameraStream&&(scannerState.cameraStream.getTracks().forEach(e=>e.stop()),scannerState.cameraStream=null),document.getElementById("scanner-video").srcObject=null
}
function switchCamera() {
  scannerState.facingMode="environment"===scannerState.facingMode?"user":"environment",startCamera()
}
function captureAndScan() {
  const e=document.getElementById("scanner-video");
  if(!e.srcObject)return void showToast("الكاميرا غير مفعّلة","error");
  const t=document.createElement("canvas");
  t.width=e.videoWidth||640,t.height=e.videoHeight||480;
  const n=t.getContext("2d");
  n.drawImage(e,0,0,t.width,t.height),processOMRImage(n.getImageData(0,0,t.width,t.height),t.width,t.height)
}
function scanFromFile(e) {
  const t=e.target.files[0];
  if(!t)return;
  const n=new Image;
  n.onload=()=> {
    const e=document.createElement("canvas");
    e.width=n.width,e.height=n.height;
    const t=e.getContext("2d");
    t.drawImage(n,0,0),processOMRImage(t.getImageData(0,0,e.width,e.height),e.width,e.height)
  },n.src=URL.createObjectURL(t)
}
function processOMRImage(e,t,n) {
  const o=scannerState.selectedModel||scannerState.selectedSubject?.models[0];
  o?(document.getElementById("scan-result-content").innerHTML='<div style="color:var(--accent);">⏳ جارٍ تحليل الصورة...</div>',setTimeout(()=> {
    try {
      const r=analyzeOMRBubbles(e,t,n,o.questions.length);
      scannerState.currentScannedAnswers=r;
      const a=[];
      Object.entries(r).forEach(([e,t])=> {
        null!==t&&"MULTI"!==t||a.push(parseInt(e))
      }),scannerState.pendingManualQuestions=a,showScanResults(r,o),a.length>0?showManualOverride(a,o):computeAndShowScore(r,o)
    }
    catch(e) {
      document.getElementById("scan-result-content").innerHTML=`<div style="color:var(--danger);">❌ خطأ في تحليل الصورة: ${e.message}</div>`,console.error("OMR error:",e)
    }
  },100)):showToast("يرجى اختيار النموذج أولاً","error")
}
function analyzeOMRBubbles(e,t,n,o) {
  const r=e.data,a= {

  },s=new Uint8Array(t*n);
  for(let e=0; e<t*n; e++) {
    const t=r[4*e],n=r[4*e+1],o=r[4*e+2];
    s[e]=Math.round(.299*t+.587*n+.114*o)
  }
  const l=new Float32Array(n);
  for(let e=0; e<n; e++) {
    let n=0;
    for(let o=0; o<t; o++)s[e*t+o]<80&&n++;
    l[e]=n/t
  }
  const i=new Float32Array(t);
  for(let e=0; e<t; e++) {
    let o=0;
    for(let r=0; r<n; r++)s[r*t+e]<80&&o++;
    i[e]=o/n
  }
  let d=Math.floor(.05*n),c=Math.floor(.95*n);
  for(let e=Math.floor(.05*n); e<Math.floor(.95*n); e++)if(l[e]>=.12) {
    d=e;
    break
  }
  for(let e=Math.floor(.95*n); e>d; e--)if(l[e]>=.12) {
    c=e;
    break
  }
  let p=Math.floor(.02*t),u=Math.floor(.98*t);
  for(let e=Math.floor(.02*t); e<Math.floor(.98*t); e++)if(i[e]>=.08) {
    p=e;
    break
  }
  for(let e=Math.floor(.98*t); e>p; e--)if(i[e]>=.08) {
    u=e;
    break
  }
  const m=(c-d)/5,g=(u-p)/(o+1),f=["A","B","C","D"];
  for(let e=0; e<o; e++) {
    const o=e+1,r=Math.floor(p+o*g),l=Math.floor(p+(o+1)*g),i=[];
    for(let e=0; e<4; e++) {
      const o=e+1,a=Math.floor(d+o*m),c=Math.floor(d+(o+1)*m),p=Math.max(2,Math.floor(.2*m)),u=Math.max(2,Math.floor(.2*g));
      let f=0,b=0;
      for(let e=a+p; e<c-p; e++)for(let o=r+u; o<l-u; o++)o>=0&&o<t&&e>=0&&e<n&&(s[e*t+o]<110&&f++,b++);
      i.push(b>0?f/b:0)
    }
    const c=i.map((e,t)=>( {
      label:f[t],ratio:e
    })).filter(e=>e.ratio>=.15);
    if(0===c.length)a[e+1]=null;
    else if(1===c.length)a[e+1]=c[0].label;
    else {
      c.sort((e,t)=>t.ratio-e.ratio);
      const t=c[0].ratio-c[1].ratio;
      a[e+1]=t<.07?"MULTI":c[0].label
    }
  }
  return a
}
function showScanResults(e,t) {
  const n=t.questions.length,o=Object.values(e).filter(e=>e&&"MULTI"!==e).length,r=Object.values(e).filter(e=>null===e||"MULTI"===e).length;
  let a=`<div style="margin-bottom:10px;font-size:12px;color:var(--text2);">\n    تم اكتشاف <strong style="color:var(--accent3);">${o}</strong> إجابة من أصل ${n}\n    ${r>0?`| <strong style="color:var(--warn);">$ {
    r
  }
  </strong> غير واضحة`:""}\n  </div>`;
  a+='<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:4px;">';
  for(let t=1; t<=n; t++) {
    const n=e[t];
    let o="var(--bg)",r="var(--text2)",s=n||"?";
    null===n&&(o="#ffb84f22",r="var(--warn)",s="؟"),"MULTI"===n&&(o="#ff4f6a22",r="var(--danger)",s="!!"),n&&"MULTI"!==n&&(o="#4f7cff22",r="var(--accent)"),a+=`<div style="padding:4px;background:${o};border-radius:6px;text-align:center;border:1px solid ${o};">\n      <div style="font-size:9px;color:var(--text3);">${t}</div>\n      <div style="font-size:13px;font-weight:700;color:${r};">${s}</div>\n    </div>`
  }
  a+="</div>",document.getElementById("scan-result-content").innerHTML=a
}
function showManualOverride(e,t) {
  const n=document.getElementById("manual-override-panel");
  document.getElementById("manual-override-list").innerHTML=e.map(e=> {
    const n=t.questions[e-1],o=n?n.choices.map(e=>e.label):["A","B","C","D"],r="MULTI"===scannerState.currentScannedAnswers[e];
    return`<div style="padding:8px;background:var(--bg);border-radius:8px;margin-bottom:8px;border:1px solid var(--border);">\n      <div style="font-size:12px;color:var(--text2);margin-bottom:6px;">\n        <strong style="color:var(--warn);">س${e}:</strong>\n        ${n?n.text.slice(0,60)+(n.text.length>60?"...":""):"سؤال رقم "+e}\n        ${r?'<span style="color:var(--danger);font-size:11px;"> (إجابات متعددة)</span>':'<span style="color:var(--warn);font-size:11px;"> (لم تُكتشف)</span>'}\n      </div>\n      <div style="display:flex;gap:6px;flex-wrap:wrap;">\n        ${o.map(t=>`<button onclick="setManualAnswer(${e},'${t}',this)"\n                  style="padding:6px 14px;border-radius:6px;border:1px solid var(--border2);\n                         background:var(--bg3);color:var(--text);cursor:pointer;\n                         font-family:Cairo,sans-serif;font-weight:600;font-size:13px;\n                         transition:all .2s;"\n                  data-q="${e}" data-label="${t}">$ {
      t
    }
    </button>`).join("")}\n        <button onclick="setManualAnswer(${e},'SKIP',this)"\n                style="padding:6px 10px;border-radius:6px;border:1px solid #ffb84f44;\n                       background:#ffb84f11;color:var(--warn);cursor:pointer;\n                       font-family:Cairo,sans-serif;font-weight:600;font-size:12px;"\n                data-q="${e}" data-label="SKIP">تجاهل</button>\n      </div>\n    </div>`
  }).join(""),n.style.display="block"
}
function setManualAnswer(e,t,n) {
  n.parentElement.querySelectorAll("button").forEach(e=> {
    e.style.background="var(--bg3)",e.style.color="var(--text)",e.style.borderColor="var(--border2)"
  }),n.style.background="SKIP"===t?"#ffb84f33":"var(--accent)",n.style.color="SKIP"===t?"var(--warn)":"#fff",n.style.borderColor="SKIP"===t?"var(--warn)":"var(--accent)",scannerState.currentScannedAnswers[e]="SKIP"===t?null:t
}
function confirmManualAnswers() {
  if(scannerState.pendingManualQuestions.filter(e=> {
    const t=scannerState.currentScannedAnswers[e];
    return null===t||"MULTI"===t
  }).length>0) {
    const e=scannerState.pendingManualQuestions.filter(e=>"MULTI"===scannerState.currentScannedAnswers[e]);
    if(e.length>0)return void showToast(`يرجى تحديد إجابة للأسئلة: ${e.join(", ")}`,"error")
  }
  const e=scannerState.selectedModel||scannerState.selectedSubject?.models[scannerState.selectedModelIdx||0];
  e&&(document.getElementById("manual-override-panel").style.display="none",computeAndShowScore(scannerState.currentScannedAnswers,e))
}
function computeAndShowScore(e,t) {
  let n=0;
  const o=t.questions.length;
  t.questions.forEach((t,o)=> {
    const r=e[o+1];
    r&&r===t.correctLabel&&n++
  });
  const r=document.getElementById("score-summary-panel");
  document.getElementById("score-display").textContent=`${n} / ${o}`,document.getElementById("score-percent").textContent=`${Math.round(n/o*100)}%`,r.style.display="block",scannerState._pendingResult= {
    studentId:scannerState.currentStudentId,modelName:t.name,modelLetter:t.letter,subject:scannerState.selectedSubject?.name||"",answers: {
      ...e
    },correct:n,total:o,score:n,percent:Math.round(n/o*100),timestamp:(new Date).toISOString()
  }
}
function saveAndNextStudent() {
  scannerState._pendingResult&&(scannerState.allResults=scannerState.allResults.filter(e=>!(e.studentId===scannerState._pendingResult.studentId&&e.subject===scannerState._pendingResult.subject)),scannerState.allResults.push(scannerState._pendingResult),saveScanResults(scannerState.allResults),renderResultsTable(),showToast(`تم حفظ نتيجة الطالب ${scannerState._pendingResult.studentId} ✓`,"success"),scannerState._pendingResult=null,scannerState.currentStudentId=null,scannerState.currentScannedAnswers=null,scannerState.pendingManualQuestions=[],document.getElementById("scanner-current-student").textContent="-",document.getElementById("scanner-student-id").value="",document.getElementById("scanner-camera-card").style.display="none",document.getElementById("scan-result-content").textContent="اضغط على زر المسح لبدء العملية...",document.getElementById("manual-override-panel").style.display="none",document.getElementById("score-summary-panel").style.display="none",stopCamera(),setTimeout(()=>document.getElementById("scanner-student-id").focus(),100))
}
function rescanCurrent() {
  scannerState.currentScannedAnswers=null,scannerState.pendingManualQuestions=[],scannerState._pendingResult=null,document.getElementById("scan-result-content").textContent="اضغط على زر المسح لبدء العملية...",document.getElementById("manual-override-panel").style.display="none",document.getElementById("score-summary-panel").style.display="none",scannerState.cameraStream||startCamera()
}
function renderResultsTable() {
  const e=document.getElementById("scanner-results-card"),t=document.getElementById("results-tbody"),n=document.getElementById("results-stats");
  if(!t)return;
  const o=scannerState.selectedSubject?.name||"",r=o?scannerState.allResults.filter(e=>e.subject===o):scannerState.allResults;
  if(0===r.length)return void(e.style.display="none");
  e.style.display="block",t.innerHTML=r.map((e,t)=> {
    const n=e.percent??Math.round(e.correct/e.total*100),o=n>=60?"var(--accent3)":n>=50?"var(--warn)":"var(--danger)";
    return`<tr>\n      <td>${t+1}</td>\n      <td><strong style="color:var(--accent);">${e.studentId}</strong></td>\n      <td>${e.modelName||e.modelLetter||"-"}</td>\n      <td>${e.correct} / ${e.total}</td>\n      <td><strong style="color:${o};">${e.score}</strong></td>\n      <td><span style="color:${o};">${n}%</span></td>\n      <td>\n        <button onclick="deleteResult('${e.studentId}','${e.subject}')"\n                style="padding:4px 10px;background:#ff4f6a22;color:var(--danger);border:none;\n                       border-radius:4px;cursor:pointer;font-family:Cairo,sans-serif;font-size:11px;">حذف</button>\n      </td>\n    </tr>`
  }).join("");
  const a=r.length>0?Math.round(r.reduce((e,t)=>e+(t.percent??0),0)/r.length):0,s=r.filter(e=>(e.percent??0)>=60).length;
  n.innerHTML=`إجمالي الطلاب: <strong>${r.length}</strong> | متوسط الدرجات: <strong style="color:var(--accent);">${a}%</strong> | الناجحون (≥60%): <strong style="color:var(--accent3);">${s}</strong>`
}
function deleteResult(e,t) {
  confirm(`حذف نتيجة الطالب ${e}؟`)&&(scannerState.allResults=scannerState.allResults.filter(n=>!(n.studentId===e&&n.subject===t)),saveScanResults(scannerState.allResults),renderResultsTable(),showToast("تم حذف النتيجة","success"))
}
function clearAllResults() {
  if(!confirm("هل أنت متأكد من حذف جميع النتائج؟ لا يمكن التراجع عن هذا الإجراء."))return;
  const e=scannerState.selectedSubject?.name||"";
  scannerState.allResults=e?scannerState.allResults.filter(t=>t.subject!==e):[],saveScanResults(scannerState.allResults),renderResultsTable(),showToast("تم مسح النتائج","success")
}
function exportResultsExcel() {
  const e=scannerState.selectedSubject?.name||"",t=e?scannerState.allResults.filter(t=>t.subject===e):scannerState.allResults;
  if(0===t.length)return void showToast("لا توجد نتائج للتصدير","error");
  const n=XLSX.utils.book_new(),o=[["الرقم الجامعي","النموذج","المادة","الإجابات الصحيحة","مجموع الأسئلة","العلامة","النسبة المئوية","تاريخ المسح"]];
  t.forEach(e=> {
    const t=e.percent??Math.round(e.correct/e.total*100);
    o.push([e.studentId,e.modelName||e.modelLetter||"",e.subject,e.correct,e.total,e.score,t+"%",e.timestamp?new Date(e.timestamp).toLocaleString("ar-EG"):""])
  });
  const r=XLSX.utils.aoa_to_sheet(o);
  r["!cols"]=[ {
    wch:18
  }, {
    wch:14
  }, {
    wch:20
  }, {
    wch:16
  }, {
    wch:14
  }, {
    wch:10
  }, {
    wch:14
  }, {
    wch:22
  }
  ],XLSX.utils.book_append_sheet(n,r,"نتائج الامتحان");
  const a=`نتائج_${e||"الامتحان"}_${(new Date).toLocaleDateString("ar-EG").replace(/\//g,"-")}.xlsx`;
  XLSX.writeFile(n,a),showToast("تم تحميل ملف Excel ✓","success")
}
document.addEventListener("DOMContentLoaded",()=> {
  initDropZone(),initColorPickers()
});