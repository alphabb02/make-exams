// JavaScript code using docx.bundle.js
// Import the library (in Node.js, you could also require it)
// Assuming docx.bundle.js has been included in your HTML or bundled with a tool like Webpack

const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell } = window.docx || require("docx"); 

// Create a new document
const doc = new Document({
    sections: [
        {
            properties: {},
            children: [
                // Adding a paragraph
                new Paragraph({
                    children: [
                        new TextRun("Hello World! "),
                        new TextRun({
                            text: "This text is bold and red.",
                            bold: true,
                            color: "FF0000"
                        })
                    ]
                }),

                // Adding another paragraph with italic text
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "This is italic text.",
                            italics: true
                        })
                    ]
                }),

                // Adding a table
                new Table({
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Cell 1")] }),
                                new TableCell({ children: [new Paragraph("Cell 2")] })
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Cell 3")] }),
                                new TableCell({ children: [new Paragraph("Cell 4")] })
                            ]
                        })
                    ]
                })
            ]
        }
    ]
});

// Generate the DOCX file and trigger download (for browser)
Packer.toBlob(doc).then((blob) => {
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "example.docx";
    link.click();
});

// Usage example for Node.js:
// const fs = require("fs");
// Packer.toBuffer(doc).then((buffer) => fs.writeFileSync("example.docx", buffer));
