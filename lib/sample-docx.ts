import JSZip from "jszip";

export const SAMPLE_OUTPUT =
  "in the case of each <b>ABR Loan</b> that is an <u>Initial</u><u>Term</u> <u>B Loan</u>, (x) prior to the Amendment No. 1 Effective Date, 2.25<i>% per annum </i>and (y) on and after the Amendment No. 1 Effective Date, (I) prior to the delivery of <u>Section 9.1</u> Financials and the related Officer\u2019s Certificate for the first full fiscal quarter following the Amendment No. 1 Effective Date pursuant to <u>Section 9.1</u>, 1.25% <i>per annum</i>";

const CONTENT_TYPES: Record<string, string> = {
  "word/document.xml":
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  "word/styles.xml":
    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
  "word/header1.xml":
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
  "word/footer1.xml":
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
  "word/footnotes.xml":
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
  "word/endnotes.xml":
    "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
  "word/comments.xml":
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
};

export function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

export function runXml(text: string, rPr = ""): string {
  const preserve = /^\s|\s$/.test(text) ? ` xml:space="preserve"` : "";
  const props = rPr ? `<w:rPr>${rPr}</w:rPr>` : "";
  return `<w:r>${props}<w:t${preserve}>${escapeXml(text)}</w:t></w:r>`;
}

export function paragraphXml(inner: string): string {
  return `<w:p>${inner}</w:p>`;
}

export function documentXml(bodyInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    ${bodyInner}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

export function defaultStylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="character" w:styleId="Emphasis">
    <w:name w:val="Emphasis"/>
    <w:rPr>
      <w:i/>
      <w:iCs/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Strong">
    <w:name w:val="Strong"/>
    <w:rPr>
      <w:b/>
      <w:bCs/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:rPr>
      <w:color w:val="0563C1"/>
      <w:u w:val="single"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr>
      <w:b/>
      <w:bCs/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="DefinedTerm">
    <w:name w:val="Defined Term"/>
    <w:basedOn w:val="Emphasis"/>
    <w:rPr>
      <w:u w:val="single"/>
    </w:rPr>
  </w:style>
</w:styles>`;
}

export function sampleDocumentXml(): string {
  const runs = [
    runXml("in the case of each ", ""),
    runXml("ABR Loan", "<w:b/><w:bCs/>"),
    runXml(" that is an ", ""),
    runXml("Initial", `<w:u w:val="single"/>`),
    runXml("Term", `<w:u w:val="single"/>`),
    runXml(" ", ""),
    runXml("B Loan", `<w:u w:val="single"/>`),
    runXml(", (x) prior to the Amendment No. 1 Effective Date, 2.25"),
    runXml("% per annum ", "<w:i/><w:iCs/>"),
    runXml(
      "and (y) on and after the Amendment No. 1 Effective Date, (I) prior to the delivery of ",
    ),
    runXml("Section 9.1", `<w:u w:val="single"/>`),
    runXml(
      " Financials and the related Officer\u2019s Certificate for the first full fiscal quarter following the Amendment No. 1 Effective Date pursuant to ",
    ),
    runXml("Section 9.1", `<w:u w:val="single"/>`),
    runXml(", 1.25% "),
    runXml("per annum", "<w:i/><w:iCs/>"),
  ];
  return documentXml(paragraphXml(runs.join("")));
}

function contentTypesXml(parts: string[]): string {
  const overrides = parts
    .filter((part) => CONTENT_TYPES[part])
    .map(
      (part) =>
        `<Override PartName="/${part}" ContentType="${CONTENT_TYPES[part]}"/>`,
    )
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${overrides}
</Types>`;
}

export async function buildDocx(options: {
  documentXml: string;
  stylesXml?: string;
  files?: Record<string, string>;
}): Promise<Uint8Array> {
  const styles = options.stylesXml ?? defaultStylesXml();
  const files = options.files ?? {};
  const zip = new JSZip();
  const parts = ["word/document.xml", "word/styles.xml", ...Object.keys(files)];
  zip.file("[Content_Types].xml", contentTypesXml(parts));
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
  );
  zip.file("word/document.xml", options.documentXml);
  zip.file("word/styles.xml", styles);
  for (const [path, xml] of Object.entries(files)) zip.file(path, xml);
  return zip.generateAsync({ type: "uint8array" });
}

export function buildSampleDocx(): Promise<Uint8Array> {
  return buildDocx({ documentXml: sampleDocumentXml() });
}
