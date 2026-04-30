let selected = null;

function addBlock(type){

    let div = document.createElement("div");
    div.className = "block";
    div.onclick = () => selectBlock(div);

    if(type === "title"){
        div.innerHTML = "عنوان جديد";
        div.style.fontSize = "22px";
        div.style.fontWeight = "bold";
    }

    if(type === "text"){
        div.innerHTML = "نص قابل للتعديل";
    }

    if(type === "line"){
        div.innerHTML = "<hr>";
    }

    if(type === "questions"){
        div.innerHTML = "{{questions}}";
    }

    div.contentEditable = true;

    document.getElementById("canvas").appendChild(div);
}

function selectBlock(el){
    selected = el;
    document.querySelectorAll(".block").forEach(b=>b.style.border="1px dashed #ccc");
    el.style.border = "2px solid red";
}

function applyStyle(){
    if(!selected) return;

    selected.style.fontSize = document.getElementById("fontSize").value + "px";
    selected.style.color = document.getElementById("color").value;
}

/* ===== PDF ===== */
function exportPDF(){
    html2pdf().from(document.getElementById("canvas")).save();
}