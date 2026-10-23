// Prints the text lines pdf.js extracts from a PDF, grouped by row (dev helper for the Premixes parser).
const path = require("node:path");

async function pdfLines(file){
  const pdfjs = require("pdfjs-dist/legacy/build/pdf.js");
  const doc = await pdfjs.getDocument({ data: new Uint8Array(require("node:fs").readFileSync(file)) }).promise;
  const lines = [];
  for (let p = 1; p <= doc.numPages; p++){
    const items = (await (await doc.getPage(p)).getTextContent()).items.filter(i => i.str.trim());
    const rows = new Map();
    items.forEach(i => { const y = Math.round(i.transform[5]); (rows.get(y) || rows.set(y, []).get(y)).push(i); });
    [...rows.entries()].sort((a, b) => b[0] - a[0])
      .forEach(([, r]) => lines.push(r.sort((a, b) => a.transform[4] - b.transform[4]).map(i => i.str.trim()).join(" ")));
  }
  return lines;
}

if (require.main === module) pdfLines(path.resolve(process.argv[2])).then(l => console.log(l.join("\n")));
module.exports = { pdfLines };
