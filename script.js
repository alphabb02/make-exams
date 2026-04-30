let questions = [];

let quill = new Quill('#editor', {
    theme: 'snow',
    placeholder: "اكتب قالب الامتحان هنا... واستخدم {{questions}}"
});

/* ================= Excel ================= */
document.getElementById("excel").addEventListener("change", e => {
    let reader = new FileReader();

    reader.onload = ev => {
        let wb = XLSX.read(new Uint8Array(ev.target.result), {type:'array'});
        let sheet = wb.Sheets[wb.SheetNames[0]];
        questions = XLSX.utils.sheet_to_json(sheet);
        alert("تم تحميل " + questions.length + " سؤال");
    };

    reader.readAsArrayBuffer(e.target.files[0]);
});

/* ================= Utils ================= */
function shuffle(a){
    return a.sort(()=>Math.random()-0.5);
}

function getDesign(){
    return {
        fontSize: document.getElementById("fontSize").value,
        lineHeight: document.getElementById("lineHeight").value,
        spacing: document.getElementById("qSpacing").value,
        layout: document.getElementById("optionLayout").value,
        columns: document.getElementById("columns").value,
        showAnswers: document.getElementById("showAnswers").checked
    };
}

/* ================= Generate ================= */
function generate(){

    if(!questions.length){
        alert("ارفع ملف Excel أولاً");
        return;
    }

    let qCount = +document.getElementById("qCount").value;
    let fCount = +document.getElementById("formCount").value;

    let template = quill.root.innerHTML;
    let design = getDesign();

    let finalHTML = "";

    for(let f=0; f<fCount; f++){

        let qs = shuffle([...questions]).slice(0, qCount);

        let content = `<div style="
            font-size:${design.fontSize}px;
            line-height:${design.lineHeight};
            column-count:${design.columns};
        ">`;

        content += `<h3>نموذج ${String.fromCharCode(65+f)}</h3>`;

        qs.forEach((q,i)=>{

            content += `<div style="margin-bottom:${design.spacing}px">`;

            content += `<p><b>${i+1}) ${q.Question}</b></p>`;

            if(q.Type === "TF"){
                content += `<p>أ) صح &nbsp;&nbsp; ب) خطأ</p>`;
            } else {
                let ops = shuffle([q.A,q.B,q.C,q.D]);

                if(design.layout === "horizontal"){
                    content += `<div style="display:flex; gap:20px;">`;
                }

                ops.forEach((o,j)=>{
                    content += `<span>${String.fromCharCode(65+j)}) ${o}</span><br>`;
                });

                if(design.layout === "horizontal"){
                    content += `</div>`;
                }
            }

            if(design.showAnswers){
                content += `<p style="color:red">الإجابة: ${q.Answer}</p>`;
            }

            content += `</div>`;
        });

        content += `</div>`;

        let page = template
            .replace("{{questions}}", content)
            .replace("{{form}}", String.fromCharCode(65+f))
            .replace("{{date}}", new Date().toLocaleDateString());

        finalHTML += page + "<div style='page-break-after:always'></div>";
    }

    document.getElementById("preview").innerHTML = finalHTML;
}

/* ================= PDF ================= */
function exportPDF(){
    html2pdf().from(document.getElementById("preview")).save("exam.pdf");
}

/* ================= Word ================= */
async function exportWord(){

    const { Document, Packer, Paragraph } = window.docx;

    let text = document.getElementById("preview").innerText.split("\n");

    let doc = new Document({
        sections:[{
            children: text.map(line => new Paragraph(line))
        }]
    });

    let blob = await Packer.toBlob(doc);
    saveAs(blob, "exam.docx");
}