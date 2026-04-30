let questions = [];
let logoBase64 = null;

document.getElementById("fileInput").addEventListener("change", readExcel);
document.getElementById("logoInput").addEventListener("change", readLogo);

function readExcel(e) {
    let file = e.target.files[0];
    let reader = new FileReader();

    reader.onload = function(e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, { type: 'array' });

        let sheet = workbook.Sheets[workbook.SheetNames[0]];
        questions = XLSX.utils.sheet_to_json(sheet);

        alert("تم تحميل " + questions.length + " سؤال");
    };

    reader.readAsArrayBuffer(file);
}

function readLogo(e) {
    let file = e.target.files[0];
    let reader = new FileReader();

    reader.onload = () => logoBase64 = reader.result;
    reader.readAsDataURL(file);
}

function shuffle(arr) {
    return arr.sort(() => Math.random() - 0.5);
}

function generateExams() {

    const { jsPDF } = window.jspdf;

    let qCount = +document.getElementById("questionCount").value;
    let fCount = +document.getElementById("formCount").value;

    let school = document.getElementById("schoolName").value;
    let subject = document.getElementById("subjectName").value;

    if (!questions.length) return alert("ارفع ملف Excel");

    let doc = new jsPDF();

    for (let f = 0; f < fCount; f++) {

        let formName = String.fromCharCode(65 + f);
        let y = 20;

        // Header
        if (logoBase64) doc.addImage(logoBase64, 'PNG', 10, 5, 20, 20);

        doc.setFontSize(14);
        doc.text(school, 105, 10, { align: "center" });
        doc.text(subject, 105, 18, { align: "center" });
        doc.text("نموذج " + formName, 180, 10);

        let selected = shuffle([...questions]).slice(0, qCount);
        let answers = [];

        selected.forEach((q, i) => {

            let type = q["Type"] || "MCQ";

            doc.text(`${i+1}) ${q["Question"]}`, 10, y);
            y += 7;

            if (type === "MCQ") {

                let options = [q["A"], q["B"], q["C"], q["D"]];
                let mixed = shuffle([...options]);

                let correctIndex = mixed.indexOf(q["Answer"]);

                mixed.forEach((opt, j) => {
                    doc.text(`${String.fromCharCode(65+j)}) ${opt}`, 15, y);
                    y += 6;
                });

                answers.push(`${i+1}-${String.fromCharCode(65+correctIndex)}`);

            } else if (type === "TF") {

                doc.text("أ) صح", 15, y);
                doc.text("ب) خطأ", 60, y);
                y += 6;

                answers.push(`${i+1}-${q["Answer"]}`);
            }

            y += 5;

            if (y > 270) {
                doc.addPage();
                y = 20;
            }
        });

        // صفحة الإجابات
        doc.addPage();
        doc.text("نموذج الإجابة - " + formName, 10, 10);

        answers.forEach((a, i) => {
            doc.text(a, 10, 20 + (i * 6));
        });

        if (f < fCount - 1) doc.addPage();

        document.getElementById("progress").innerText =
            "تم إنشاء نموذج " + formName;
    }

    doc.save("All_Exams.pdf");
}
