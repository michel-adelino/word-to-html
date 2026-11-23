import assert from "node:assert/strict";
import { describe, it } from "node:test";
import {
  DOMParser,
  type Element as XmlElement,
} from "@xmldom/xmldom";
import JSZip from "jszip";
import { ConvertError, convertDocx } from "./convert-docx";
import {
  SAMPLE_OUTPUT,
  buildDocx,
  buildSampleDocx,
  documentXml,
  paragraphXml,
  runXml,
} from "./sample-docx";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

async function partXml(data: Uint8Array, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(data);
  const file = zip.file(path);
  assert.ok(file, `missing ${path}`);
  return file.async("string");
}

function localName(node: XmlElement): string {
  return node.localName || node.tagName.replace(/^.*:/, "");
}

function paragraphTexts(xml: string): string[] {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const texts: string[] = [];
  const walkParagraph = (paragraph: XmlElement) => {
    let text = "";
    const walk = (node: XmlElement) => {
      for (let child = node.firstChild; child; child = child.nextSibling) {
        if (child.nodeType !== 1) continue;
        const element = child as XmlElement;
        if (element.namespaceURI && element.namespaceURI !== W_NS) {
          walk(element);
          continue;
        }
        const name = localName(element);
        if (name === "p") continue;
        if (name === "t" || name === "delText") text += element.textContent ?? "";
        else walk(element);
      }
    };
    walk(paragraph);
    texts.push(text);
  };

  const walk = (node: XmlElement) => {
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== 1) continue;
      const element = child as XmlElement;
      if (localName(element) === "p" && (!element.namespaceURI || element.namespaceURI === W_NS)) {
        walkParagraph(element);
      }
      walk(element);
    }
  };
  walk(doc.documentElement as XmlElement);
  return texts.filter((text) => text.length > 0);
}

describe("convertDocx", () => {
  it("writes the sample clause as HTML tags without merging runs", async () => {
    const result = await convertDocx(await buildSampleDocx());
    assert.equal(result.preview[0]?.after, SAMPLE_OUTPUT);
    assert.deepEqual(result.stats, {
      bold: 1,
      italic: 2,
      underline: 5,
      strike: 0,
      subscript: 0,
      superscript: 0,
      runs: 8,
    });
    assert.equal(result.truncated, false);

    const stored = paragraphTexts(await partXml(result.data, "word/document.xml"));
    assert.deepEqual(stored, [SAMPLE_OUTPUT]);
    const xml = await partXml(result.data, "word/document.xml");
    assert.match(xml, /&lt;b&gt;ABR Loan&lt;\/b&gt;/);
    assert.match(xml, /w:val="0"/);
  });

  it("nests combined formatting and keeps a second pass stable", async () => {
    const source = await buildDocx({
      documentXml: documentXml(
        paragraphXml(
          [
            runXml("defined", `<w:b/><w:i/><w:u w:val="double"/>`),
            runXml("2", `<w:vertAlign w:val="subscript"/>`),
            runXml("n", `<w:vertAlign w:val="superscript"/>`),
            runXml("gone", "<w:strike/>"),
            runXml("gone too", "<w:dstrike/>"),
            runXml("a < b & c", "<w:b/>"),
          ].join(""),
        ),
      ),
    });

    const once = await convertDocx(source);
    assert.equal(
      once.preview[0]?.after,
      "<b><i><u>defined</u></i></b><sub>2</sub><sup>n</sup><s>gone</s><s>gone too</s><b>a < b & c</b>",
    );

    const twice = await convertDocx(once.data);
    assert.equal(twice.preview[0]?.after, once.preview[0]?.after);
    assert.equal(twice.stats.runs, 0);
  });

  it("uses character styles, ignores paragraph styles and hyperlinks", async () => {
    const source = await buildDocx({
      documentXml: documentXml(
        [
          paragraphXml(
            `<w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>latin</w:t></w:r>`,
          ),
          paragraphXml(
            `<w:r><w:rPr><w:rStyle w:val="DefinedTerm"/></w:rPr><w:t>Borrower</w:t></w:r>`,
          ),
          paragraphXml(
            `<w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r>`,
          ),
          paragraphXml(
            `<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Section 9.1</w:t></w:r>`,
          ),
          paragraphXml(
            `<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:b/></w:rPr><w:t>link</w:t></w:r>`,
          ),
          paragraphXml(runXml("plain", `<w:b w:val="0"/><w:i w:val="false"/>`)),
          paragraphXml(runXml("open", `<w:u w:val="none"/>`)),
        ].join(""),
      ),
    });

    const result = await convertDocx(source);
    assert.deepEqual(
      result.preview.map((paragraph) => paragraph.after),
      [
        "<i>latin</i>",
        "<i><u>Borrower</u></i>",
        "Title",
        "Section 9.1",
        "<b>link</b>",
        "plain",
        "open",
      ],
    );
  });

  it("preserves spaces, merges adjacent text nodes, and splits on tabs", async () => {
    const source = await buildDocx({
      documentXml: documentXml(
        [
          paragraphXml(
            `<w:r><w:rPr><w:b/></w:rPr><w:t>ABR</w:t><w:t xml:space="preserve"> Loan</w:t></w:r>`,
          ),
          paragraphXml(
            `<w:r><w:rPr><w:i/></w:rPr><w:t>hello</w:t><w:tab/><w:t>world</w:t></w:r>`,
          ),
          `<w:tbl><w:tr><w:tc><w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
        ].join(""),
      ),
    });

    const result = await convertDocx(source);
    assert.deepEqual(
      result.preview.map((paragraph) => paragraph.after),
      ["<b>ABR Loan</b>", "<i>hello</i>\t<i>world</i>", "<u>Cell</u>"],
    );
    assert.equal(result.preview[0]?.before[0]?.bold, true);
    assert.equal(result.preview[0]?.before[0]?.text, "ABR Loan");
  });

  it("converts headers as well as the body", async () => {
    const source = await buildDocx({
      documentXml: documentXml(paragraphXml(runXml("Body", "<w:b/>"))),
      files: {
        "word/header1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Confidential</w:t></w:r></w:p>
</w:hdr>`,
      },
    });
    const result = await convertDocx(source);
    assert.equal(result.preview[0]?.after, "<b>Body</b>");
    const header = await partXml(result.data, "word/header1.xml");
    assert.match(header, /&lt;i&gt;Confidential&lt;\/i&gt;/);
  });

  it("caps the preview and still converts the rest of the document", async () => {
    const paragraphs = Array.from({ length: 41 }, (_, index) =>
      paragraphXml(index === 40 ? runXml("last", "<w:b/>") : runXml(`P${index}`)),
    ).join("");
    const result = await convertDocx(
      await buildDocx({ documentXml: documentXml(paragraphs) }),
    );
    assert.equal(result.preview.length, 40);
    assert.equal(result.truncated, true);
    assert.equal(result.stats.bold, 1);
    const stored = paragraphTexts(await partXml(result.data, "word/document.xml"));
    assert.equal(stored.at(-1), "<b>last</b>");
  });

  it("rejects encrypted and unreadable files", async () => {
    const encrypted = new JSZip();
    encrypted.file("EncryptionInfo", "locked");
    encrypted.file("EncryptedPackage", "locked");
    await assert.rejects(
      async () => convertDocx(await encrypted.generateAsync({ type: "uint8array" })),
      (error: unknown) => {
        assert.ok(error instanceof ConvertError);
        assert.equal(error.code, "encrypted");
        assert.match(error.message, /password-protected/);
        return true;
      },
    );

    await assert.rejects(
      () => convertDocx(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0])),
      (error: unknown) => {
        assert.ok(error instanceof ConvertError);
        assert.equal(error.code, "invalid");
        return true;
      },
    );
  });
});
