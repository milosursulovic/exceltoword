const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

if (process.argv.length < 5) {
  console.log("Upotreba:");
  console.log("  node index.js <excelPath> <templatePath> <outputFolder>");
  console.log("Primer:");
  console.log("  node index.js input.xlsx template.docx out");
  process.exit(1);
}

const excelPath = process.argv[2];
const templatePath = process.argv[3];
const outputFolder = process.argv[4];

if (!fs.existsSync(excelPath)) {
  console.error("Excel fajl ne postoji:", excelPath);
  process.exit(1);
}

if (!fs.existsSync(templatePath)) {
  console.error("Word template fajl ne postoji:", templatePath);
  process.exit(1);
}

if (!fs.existsSync(outputFolder)) {
  fs.mkdirSync(outputFolder, { recursive: true });
}

const workbook = XLSX.readFile(excelPath);

const firstSheetName = workbook.SheetNames[0];
const secondSheetName = workbook.SheetNames[1];
const thirdSheetName = workbook.SheetNames[2];

const worksheet1 = workbook.Sheets[firstSheetName];
const worksheet2 = workbook.Sheets[secondSheetName];
const worksheet3 = workbook.Sheets[thirdSheetName];

const firstSheetRows = XLSX.utils.sheet_to_json(worksheet1, { defval: "" });
const secondSheetRows = XLSX.utils.sheet_to_json(worksheet2, { defval: "" });
const thirdSheetRows = XLSX.utils.sheet_to_json(worksheet3, { defval: "" });

function getDosije(row) {
  return String(row.Dosije ?? row.dosije ?? "").trim();
}

const secondByDosije = new Map();
secondSheetRows.forEach((r) => {
  const key = getDosije(r);
  if (key) {
    secondByDosije.set(key, r);
  }
});

const thirdByDosije = new Map();
thirdSheetRows.forEach((r) => {
  const key = getDosije(r);
  if (key) {
    thirdByDosije.set(key, r);
  }
});

function cleanOrgUnit(value) {
  if (!value) return "";
  return value.replace(/^[\d.]+(?=\p{L})/u, "").trim();
}

const templateBinary = fs.readFileSync(templatePath, "binary");

let generated = 0;

firstSheetRows.forEach((row, index) => {
  const ime = String(row.Ime || "").trim();
  const prezime = String(row.Prezime || "").trim();
  const strucnaSprema = String(row.StrucnaSprema || "").trim();
  const pozicija = String(row.Pozicija || "").trim();
  const radnoMesto = String(row.RadnoMesto || "").trim();
  const organizacionaJedinica = cleanOrgUnit(row.OrganizacionaJedinica);

  const dosije = getDosije(row);

  if (!dosije) {
    console.log(`Red ${index + 2}: nema Dosije, preskačem.`);
    return;
  }

  const rowSheet2 = secondByDosije.get(dosije) || {};
  const rowSheet3 = thirdByDosije.get(dosije) || {};

  const JMBG = String(rowSheet2.JMBG || "").trim();

  const data = {
    Ime: ime,
    Prezime: prezime,
    StrucnaSprema: strucnaSprema,
    Pozicija: pozicija,
    RadnoMesto: radnoMesto,
    OrganizacionaJedinica: organizacionaJedinica,
    Dosije: dosije,
    Pol:
      rowSheet2.Pol === "M"
        ? "Заспленом"
        : rowSheet2.Pol === "Ž"
        ? "Заспленој"
        : "",

    UgovorORaduBroj: String(rowSheet3.UgovorORaduBroj || "").trim(),
    UgovorORaduDatum: String(rowSheet3.UgovorORaduDatum || "").trim(),
  };

  const zip = new PizZip(templateBinary);
  const doc = new Docxtemplater(zip, {
    delimiters: { start: "[[", end: "]]" },
  });

  doc.render(data);

  const buf = doc.getZip().generate({
    type: "nodebuffer",
  });

  const safeIme = ime || "BezImena";
  const safePrezime = prezime || "BezPrezimena";
  const safeJMBG = JMBG || "BezJMBG";

  const fileName = `${safeJMBG}_${safeIme}_${safePrezime}.docx`;
  const outputPath = path.join(outputFolder, fileName);

  fs.writeFileSync(outputPath, buf);

  console.log(`Kreiran fajl: ${outputPath}`);
  generated++;
});

console.log(`\nGotovo! Ukupno generisano: ${generated} fajlova.`);
