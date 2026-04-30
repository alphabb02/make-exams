let questions = [];
let usedQuestions = new Set();

let quill = new Quill('#editor', {
    theme: 'snow',
    modules: {
        toolbar: [
            [{ font: [] }, { size: [] }],
            ['bold', 'italic', 'underline'],
            [{ color: [] }, { background: [] }],
            [{ align: [] }],
            ['image', 'code-block']
        ]
    }
});

document.getElementById("excel").addEventListener("change", e => {

    let reader = new FileReader();

    reader.onload = ev => {
        let wb = XLSX.read(new Uint8Array(ev.target.result), {type:'array'});
        let sheet = wb.Sheets[wb.SheetNames[0]];
        questions = XLSX.utils.sheet_to_json(sheet);
    };

    reader.readAsArrayBuffer(e.target.files[0]);
});

function saveTemplate(){
    localStorage.setItem("template", quill.root.innerHTML);
}

function loadTemplate(){
    quill.root.innerHTML = localStorage.getItem("template") || "";
}

function shuffle(a){
    return a.sort(()=>Math.random()-0.5);
}

function pickQuestions(count){
    let available = questions.filter(q => !usedQuestions.has(q.Question));
    let selected = shuffle(available).slice(0, count);
    selected.forEach(q => usedQuestions.add(q.Question));
    return selected;
}

function generate(){

    let qCount = +document.getElementById("qCount").value;
    let fCount = +document.getElementById("formCount").value;

    let template = quill.root.innerHTML;

    let finalHTML = "";

    for(let f=0; f<fCount; f++){

        let qs = pickQuestions(qCount);

        let content = `<h3>نموذج ${String.fromCharCode(65+f)}</h3>`;

        qs.forEach((q,i)=>{

            content += `<p>${i+1}) ${q.Question}</p>`;

            if(q.Type === "TF"){
                content += `<p>أ) صح ب) خطأ</p>`;
            }
            else{
                let ops = shuffle([q.A,q.B,q.C,q.D]);
                ops.forEach((o,j)=>{
                    content += `<p>${String.fromCharCode(65+j)}) ${o}</p>`;
                });
            }
        });

        let page = template
            .replace("{{questions}}", content)
            .replace("{{form}}", String.fromCharCode(65+f))
            .replace("{{date}}", new Date().toLocaleDateString());

        finalHTML += page + "<div style='page-break-after:always'></div>";
    }

    document.getElementById("preview").innerHTML = finalHTML;
}

function exportPDF(){
    html2pdf().from(document.getElementById("preview")).save("exam.pdf");
}

async function exportWord(){

    const { Document, Packer, Paragraph } = window.docx;

    let lines = document.getElementById("preview").innerText.split("\n");

    let doc = new Document({
        sections:[{
            children: lines.map(l=> new Paragraph(l))
        }]
    });

    let blob = await Packer.toBlob(doc);
    saveAs(blob,"exam.docx");
}