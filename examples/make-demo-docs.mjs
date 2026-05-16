// Generates the demo Office files:
//   examples/excel-demo/starter.xlsx — a blank one-sheet workbook
//   examples/word-demo/starter.docx  — the Word demo starter text with
//                                      deliberate house-style violations
//                                      and real Heading 2 sections
//
// A .docx/.xlsx is just a ZIP of OOXML parts, so we hand-write the minimal
// valid part set and package it with the system `zip`. Run from the repo
// root:  node examples/make-demo-docs.mjs
//
// (We never write Office files the add-in might have open — these are
// fresh fixtures generated on disk, not live documents.)

import { mkdtemp, mkdir, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const run = promisify(execFile);
const here = dirname(fileURLToPath(import.meta.url));

// Write a {path: contents} part map into a temp dir and zip it into `out`.
async function pack(parts, out) {
  const dir = await mkdtemp(join(tmpdir(), "ooxml-"));
  try {
    for (const [rel, body] of Object.entries(parts)) {
      const abs = join(dir, rel);
      await mkdir(dirname(abs), { recursive: true });
      await writeFile(abs, body);
    }
    await rm(out, { force: true });
    // -X drops extra file attributes; relative paths become the zip entries.
    await run("zip", ["-X", "-r", "-q", out, ".", "-i", "*"], { cwd: dir });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const ODOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

// ---------------------------------------------------------------------------
// Blank workbook
// ---------------------------------------------------------------------------
async function makeXlsx(out) {
  const W = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  await pack(
    {
      "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CT}">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
      "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL}">
<Relationship Id="rId1" Type="${ODOC}" Target="xl/workbook.xml"/>
</Relationships>`,
      "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
      "xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL}">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
      "xl/worksheets/sheet1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${W}"><sheetData/></worksheet>`,
    },
    out,
  );
}

// ---------------------------------------------------------------------------
// Word starter doc — paragraphs are [text, styleId|null].
// ---------------------------------------------------------------------------
const PARAS = [
  ["Acme relay Q3 rollout", "Title"],
  ["We are thrilled to announce that acme relay will ship to 5 pilot teams on 6/3/26!", null],
  ["Background", "Heading2"],
  ["The relay project began in March and was scoped by the platform team.", null],
  ["Next steps", "Heading2"],
  ["Teams will get onboarding docs and we will collect feedback after two weeks.", null],
];

const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

async function makeDocx(out) {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const body = PARAS.map(([text, style]) => {
    const pPr = style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : "";
    return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
  }).join("");

  await pack(
    {
      "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CT}">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`,
      "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL}">
<Relationship Id="rId1" Type="${ODOC}" Target="word/document.xml"/>
</Relationships>`,
      "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL}">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
      "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
      "word/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W}">
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
</w:styles>`,
    },
    out,
  );
}

const xlsx = join(here, "excel-demo", "starter.xlsx");
const docx = join(here, "word-demo", "starter.docx");
await makeXlsx(xlsx);
await makeDocx(docx);

// Sanity: a valid OOXML zip must contain [Content_Types].xml.
for (const f of [xlsx, docx]) {
  const { stdout } = await run("unzip", ["-l", f]);
  if (!stdout.includes("[Content_Types].xml")) {
    throw new Error(`Generated ${f} is missing [Content_Types].xml`);
  }
}
console.log(`Wrote:\n  ${xlsx}\n  ${docx}`);
