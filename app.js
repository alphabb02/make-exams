let questionBank=[];
let answerKeys={};
let uploadedHeader="";
let uploadedLogos={ministry:"",school:""};
let lockHeader=false;

let examInfo={
 school:"",subject:"",grade:"",
 duration:"",fontSize:14,align:"right"
};

/* Excel */
function loadExcel(file){
 const r=new FileReader();
 r.onload=e=>{
  const wb=XLSX.read(e.target.result,{type:"binary"});
  const sh=wb.Sheets[wb.SheetNames[0]];
  questionBank=XLSX.utils.sheet_to_json(sh);
  alert("تم تحميل بنك الأسئلة");
 };
 r.readAsBinaryString(file);
}

/* Upload Images */
function uploadLogo(f,t){
 if(lockHeader)return alert("مقفول");
 const r=new FileReader();
 r.onload=e=>uploadedLogos[t]=e.target.result;
 r.readAsDataURL(f);
}
function uploadHeader(f){
 if(lockHeader)return alert("مقفول");
 const r=new FileReader();
 r.onload=e=>uploadedHeader=e.target.result;
 r.readAsDataURL(f);
}

/* Exam Pages */
function createHeader(){
 if(uploadedHeader)
  return `<div class="header"><img src="${uploadedHeader}" style="max-width:100%"></div>`;
 return `
 <div class="header" style="display:flex;justify-content:space-between">
  ${uploadedLogos.ministry?`<img src="${uploadedLogos.ministry}" height="40">`:""}
  <div>
   <b>${examInfo.school}</b><br>
   ${examInfo.subject} - ${examInfo.grade}<br>
   الزمن:${examInfo.duration}
  </div>
  ${uploadedLogos.school?`<img src="${uploadedLogos.school}" height="40">`:""}
 </div>`;
}

function generateExams(){
 if(!questionBank.length) return alert("لا يوجد أسئلة");
 lockHeader=true;
 const preview=document.getElementById("preview");
 preview.innerHTML="";
 const models=+modelsCount.value;
 const qCount=+questionsCount.value;

 for(let m=0;m<models;m++){
  const model=String.fromCharCode(65+m);
  const qs=shuffle([...questionBank]).slice(0,qCount);
  let key=[];
  let page=createPage(model,1);
  let pNo=1;
  let content=page.querySelector(".content");
  preview.appendChild(page);

  qs.forEach((q,i)=>{
   const ans=shuffle([q.A,q.B,q.C,q.D]);
   key.push({q:i+1,a:ans.indexOf(q[q.الإجابة]) + 1});
   const div=document.createElement("div");
   div.className="question";
   div.style.fontSize=examInfo.fontSize+"px";
   div.innerHTML=`<b>${i+1})</b> ${q.السؤال}<br>`+
    ans.map((a,j)=>`(${j+1}) ${a}`).join("<br>");
   content.appendChild(div);
   if(content.scrollHeight>900){
    pNo++;
    page=createPage(model,pNo);
    content=page.querySelector(".content");
    preview.appendChild(page);
    content.appendChild(div);
   }
  });

  answerKeys[model]=key;
  html2pdf().from(page).save(`Exam_${model}.pdf`);
 }
}

function createPage(model,p){
 const d=document.createElement("div");
 d.className="exam-page";
 d.innerHTML=`
 ${createHeader()}
 <h3 style="text-align:center">نموذج ${model}</h3><hr>
 <div class="content"></div>
 <div class="footer">الصفحة ${p}</div>`;
 new Sortable(d.querySelector(".content"));
 return d;
}

/* Grading */
function openGrading(){ /* نفس الكود السابق */ }
function shuffle(a){return a.sort(()=>Math.random()-.5);}
