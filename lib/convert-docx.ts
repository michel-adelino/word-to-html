import {
  DOMParser,
  XMLSerializer,
  type Document as XmlDocument,
  type Element as XmlElement,
  type Node as XmlNode,
} from "@xmldom/xmldom";
import JSZip from "jszip";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_NS = "http://www.w3.org/XML/1998/namespace";

export const PREVIEW_PARAGRAPH_LIMIT = 40;

const INVALID_MESSAGE =
  "This file isn't a readable .docx. If it's an older .doc, save it as .docx in Word and try again.";
const ENCRYPTED_MESSAGE =
  "This document is password-protected. Remove the password in Word, then convert it again.";

export class ConvertError extends Error {
  readonly code: "encrypted" | "invalid";

  constructor(message: string, code: "encrypted" | "invalid") {
    super(message);
    this.name = "ConvertError";
    this.code = code;
  }
}

export type TextFormat = {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  subscript: boolean;
  superscript: boolean;
};

export type PreviewSpan = TextFormat & {
  text: string;
};

export type PreviewParagraph = {
  before: PreviewSpan[];
  after: string;
};

export type ConvertStats = {
  bold: number;
  italic: number;
  underline: number;
  strike: number;
  subscript: number;
  superscript: number;
  runs: number;
};

export type ConvertResult = {
  data: Uint8Array;
  stats: ConvertStats;
  preview: PreviewParagraph[];
  truncated: boolean;
};

const EMPTY_FORMAT: TextFormat = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  subscript: false,
  superscript: false,
};

type StyleInfo = {
  type: string;
  name: string;
  basedOn: string | null;
  rPr: XmlElement | null;
};

type ToggleField = "b" | "bCs" | "i" | "iCs" | "strike" | "dstrike";

type FormatState = Record<ToggleField, boolean> & {
  underline: string | null;
  vert: "sub" | "sup" | null;
};

type RawPiece =
  | { kind: "text"; nodes: XmlElement[]; text: string }
  | { kind: "gap"; text: string };

type Piece = {
  raw: string;
  after: string;
  format: TextFormat;
};

function localName(node: XmlElement): string {
  return node.localName || node.tagName.replace(/^.*:/, "");
}

function isWordElement(node: XmlElement, name: string): boolean {
  if (localName(node) !== name) return false;
  if (node.namespaceURI) return node.namespaceURI === W_NS;
  const tag = node.tagName;
  return tag === `w:${name}` || tag === name;
}

function firstChild(parent: XmlElement, name: string): XmlElement | null {
  for (let child = parent.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1 && isWordElement(child as XmlElement, name)) {
      return child as XmlElement;
    }
  }
  return null;
}

function attr(element: XmlElement, name: string): string | null {
  const namespaced = element.getAttributeNS(W_NS, name);
  if (namespaced != null) return namespaced;
  const prefixed = element.getAttribute(`w:${name}`);
  if (prefixed != null) return prefixed;
  return element.getAttribute(name);
}

function collect(root: XmlNode, name: string): XmlElement[] {
  const found: XmlElement[] = [];
  const walk = (node: XmlNode) => {
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== 1) continue;
      const element = child as XmlElement;
      if (isWordElement(element, name)) found.push(element);
      walk(element);
    }
  };
  walk(root);
  return found;
}

function parseXml(xml: string): XmlDocument {
  try {
    const doc = new DOMParser({
      onError: (level) => {
        if (level === "fatalError") {
          throw new ConvertError(INVALID_MESSAGE, "invalid");
        }
      },
    }).parseFromString(xml, "application/xml");
    if (!doc.documentElement) {
      throw new ConvertError(INVALID_MESSAGE, "invalid");
    }
    return doc;
  } catch (error) {
    if (error instanceof ConvertError) throw error;
    throw new ConvertError(INVALID_MESSAGE, "invalid");
  }
}

function serialize(doc: XmlDocument, original: string): string {
  const body = new XMLSerializer().serializeToString(doc);
  if (body.startsWith("<?xml")) return body;
  const declared = original.match(/^\uFEFF?\s*(<\?xml[\s\S]*?\?>)/);
  const declaration =
    declared?.[1] ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
  return `${declaration}${body}`;
}

function loadStyles(xml: string | null): Map<string, StyleInfo> {
  const styles = new Map<string, StyleInfo>();
  if (!xml) return styles;
  const doc = parseXml(xml);
  for (const style of collect(doc, "style")) {
    const id = attr(style, "styleId");
    if (!id) continue;
    const basedOn = firstChild(style, "basedOn");
    const name = firstChild(style, "name");
    styles.set(id, {
      type: (attr(style, "type") ?? "").toLowerCase(),
      name: name ? (attr(name, "val") ?? "") : "",
      basedOn: basedOn ? attr(basedOn, "val") : null,
      rPr: firstChild(style, "rPr"),
    });
  }
  return styles;
}

function isHyperlinkStyle(style: StyleInfo): boolean {
  const name = style.name.replace(/\s+/g, "").toLowerCase();
  return name === "hyperlink" || name === "followedhyperlink";
}

function toggleState(element: XmlElement | null): "absent" | "on" | "off" {
  if (!element) return "absent";
  const value = attr(element, "val");
  if (value == null || value === "") return "on";
  const normalized = value.toLowerCase();
  if (normalized === "0" || normalized === "false" || normalized === "off") {
    return "off";
  }
  return "on";
}

function applyToggle(current: boolean, state: "absent" | "on" | "off"): boolean {
  if (state === "absent") return current;
  if (state === "off") return false;
  return !current;
}

function emptyState(): FormatState {
  return {
    b: false,
    bCs: false,
    i: false,
    iCs: false,
    strike: false,
    dstrike: false,
    underline: null,
    vert: null,
  };
}

function applyRPr(state: FormatState, rPr: XmlElement | null) {
  if (!rPr) return;
  const fields: ToggleField[] = ["b", "bCs", "i", "iCs", "strike", "dstrike"];
  for (const field of fields) {
    state[field] = applyToggle(state[field], toggleState(firstChild(rPr, field)));
  }

  const underline = firstChild(rPr, "u");
  if (underline) {
    const value = attr(underline, "val");
    state.underline = value == null || value === "" ? "single" : value.toLowerCase();
  }

  const align = firstChild(rPr, "vertAlign");
  if (align) {
    const value = (attr(align, "val") ?? "").toLowerCase();
    if (value === "subscript") state.vert = "sub";
    else if (value === "superscript") state.vert = "sup";
    else state.vert = null;
  }
}

function resolveFormat(run: XmlElement, styles: Map<string, StyleInfo>): TextFormat {
  const rPr = firstChild(run, "rPr");
  const styleId = rPr ? firstChild(rPr, "rStyle") : null;
  const chain: StyleInfo[] = [];
  const seen = new Set<string>();
  let id = styleId ? attr(styleId, "val") : null;

  while (id && !seen.has(id)) {
    const style = styles.get(id);
    if (!style || (style.type && style.type !== "character")) break;
    seen.add(id);
    chain.push(style);
    id = style.basedOn;
  }
  chain.reverse();

  const state = emptyState();
  for (const style of chain) {
    if (isHyperlinkStyle(style)) continue;
    applyRPr(state, style.rPr);
  }
  applyRPr(state, rPr);

  const underline = state.underline;
  const underlined =
    underline != null &&
    underline !== "none" &&
    underline !== "0" &&
    underline !== "false" &&
    underline !== "off";

  return {
    bold: state.b || state.bCs,
    italic: state.i || state.iCs,
    underline: underlined,
    strike: state.strike || state.dstrike,
    subscript: state.vert === "sub",
    superscript: state.vert === "sup",
  };
}

function hasFormat(format: TextFormat): boolean {
  return (
    format.bold ||
    format.italic ||
    format.underline ||
    format.strike ||
    format.subscript ||
    format.superscript
  );
}

export function wrapFormattedText(text: string, format: TextFormat): string {
  let wrapped = text;
  if (format.superscript) wrapped = `<sup>${wrapped}</sup>`;
  else if (format.subscript) wrapped = `<sub>${wrapped}</sub>`;
  if (format.strike) wrapped = `<s>${wrapped}</s>`;
  if (format.underline) wrapped = `<u>${wrapped}</u>`;
  if (format.italic) wrapped = `<i>${wrapped}</i>`;
  if (format.bold) wrapped = `<b>${wrapped}</b>`;
  return wrapped;
}

function rawPieces(run: XmlElement): RawPiece[] {
  const pieces: RawPiece[] = [];
  let pending: XmlElement[] | null = null;
  let pendingKind: "t" | "delText" | null = null;

  const flush = () => {
    if (!pending || !pendingKind) return;
    pieces.push({
      kind: "text",
      nodes: pending,
      text: pending.map((node) => node.textContent ?? "").join(""),
    });
    pending = null;
    pendingKind = null;
  };

  for (let child = run.firstChild; child; child = child.nextSibling) {
    if (child.nodeType !== 1) continue;
    const element = child as XmlElement;
    if (!isWordElement(element, localName(element))) {
      flush();
      continue;
    }
    const name = localName(element);
    if (name === "t" || name === "delText") {
      if (pending && pendingKind === name) pending.push(element);
      else {
        flush();
        pending = [element];
        pendingKind = name;
      }
      continue;
    }
    flush();
    if (name === "tab" || name === "ptab") pieces.push({ kind: "gap", text: "\t" });
    else if (name === "br" || name === "cr") pieces.push({ kind: "gap", text: "\n" });
  }
  flush();
  return pieces;
}

function hasPreserve(element: XmlElement): boolean {
  return (
    element.getAttributeNS(XML_NS, "space") === "preserve" ||
    element.getAttribute("xml:space") === "preserve"
  );
}

function writeText(element: XmlElement, text: string, preserve: boolean) {
  element.textContent = text;
  if (preserve || /^\s|\s$/.test(text)) {
    element.setAttributeNS(XML_NS, "xml:space", "preserve");
  }
}

function setWordVal(element: XmlElement, value: string) {
  const names: string[] = [];
  for (let index = 0; index < element.attributes.length; index += 1) {
    const attribute = element.attributes.item(index);
    if (!attribute) continue;
    const name = attribute.localName || attribute.name;
    if (name === "val") names.push(attribute.name);
  }
  for (const name of names) element.removeAttribute(name);
  element.setAttributeNS(W_NS, "w:val", value);
}

function ensureChild(parent: XmlElement, name: string): XmlElement {
  const existing = firstChild(parent, name);
  if (existing) return existing;
  const prefix = parent.prefix || parent.tagName.split(":")[0] || "w";
  const qualified = prefix && parent.tagName.includes(":") ? `${prefix}:${name}` : `w:${name}`;
  const created = parent.ownerDocument!.createElementNS(W_NS, qualified);
  parent.appendChild(created);
  return created;
}

function ensureRPr(run: XmlElement): XmlElement {
  const existing = firstChild(run, "rPr");
  if (existing) return existing;
  const prefix = run.prefix || "w";
  const created = run.ownerDocument!.createElementNS(W_NS, `${prefix}:rPr`);
  run.insertBefore(created, run.firstChild);
  return created;
}

function disableFormat(run: XmlElement, format: TextFormat) {
  const rPr = ensureRPr(run);
  if (format.bold) {
    setWordVal(ensureChild(rPr, "b"), "0");
    setWordVal(ensureChild(rPr, "bCs"), "0");
  }
  if (format.italic) {
    setWordVal(ensureChild(rPr, "i"), "0");
    setWordVal(ensureChild(rPr, "iCs"), "0");
  }
  if (format.strike) {
    setWordVal(ensureChild(rPr, "strike"), "0");
    setWordVal(ensureChild(rPr, "dstrike"), "0");
  }
  if (format.underline) setWordVal(ensureChild(rPr, "u"), "none");
  if (format.subscript || format.superscript) {
    setWordVal(ensureChild(rPr, "vertAlign"), "baseline");
  }
}

function bump(stats: ConvertStats, format: TextFormat) {
  stats.runs += 1;
  if (format.bold) stats.bold += 1;
  if (format.italic) stats.italic += 1;
  if (format.underline) stats.underline += 1;
  if (format.strike) stats.strike += 1;
  if (format.subscript) stats.subscript += 1;
  if (format.superscript) stats.superscript += 1;
}

function processRun(
  run: XmlElement,
  styles: Map<string, StyleInfo>,
  stats: ConvertStats,
): Piece[] {
  const format = resolveFormat(run, styles);
  const formatted = hasFormat(format);
  const pieces = rawPieces(run);
  const out: Piece[] = [];
  let tagged = false;

  for (const piece of pieces) {
    if (piece.kind === "gap") {
      out.push({ raw: piece.text, after: piece.text, format: EMPTY_FORMAT });
      continue;
    }
    if (piece.text === "") continue;
    if (!formatted) {
      out.push({ raw: piece.text, after: piece.text, format: EMPTY_FORMAT });
      continue;
    }
    const after = wrapFormattedText(piece.text, format);
    const preserve =
      piece.nodes.some(hasPreserve) || /^\s|\s$/.test(piece.text);
    writeText(piece.nodes[0], after, preserve);
    for (const extra of piece.nodes.slice(1)) {
      extra.parentNode?.removeChild(extra);
    }
    tagged = true;
    out.push({ raw: piece.text, after, format });
  }

  if (tagged) {
    disableFormat(run, format);
    bump(stats, format);
  }
  return out;
}

function runsInParagraph(paragraph: XmlElement): XmlElement[] {
  const runs: XmlElement[] = [];
  const walk = (node: XmlElement) => {
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== 1) continue;
      const element = child as XmlElement;
      if (isWordElement(element, "p")) continue;
      if (isWordElement(element, "r")) runs.push(element);
      else walk(element);
    }
  };
  walk(paragraph);
  return runs;
}

function shouldProcessPart(path: string): boolean {
  const normalized = path.replace(/^\/+/, "");
  if (!normalized.startsWith("word/") || !normalized.endsWith(".xml")) return false;
  if (normalized.includes("/_rels/")) return false;
  if (/^word\/theme\//i.test(normalized)) return false;
  if (
    /^word\/(styles|stylesWithEffects|settings|fontTable|webSettings|numbering)\.xml$/i.test(
      normalized,
    )
  ) {
    return false;
  }
  return true;
}

function transformPart(
  xml: string,
  styles: Map<string, StyleInfo>,
  stats: ConvertStats,
  capturePreview: boolean,
): { xml: string; changed: boolean; preview: PreviewParagraph[]; truncated: boolean } {
  const doc = parseXml(xml);
  const preview: PreviewParagraph[] = [];
  let truncated = false;
  let changed = false;
  const seen = new Set<XmlElement>();

  const consume = (run: XmlElement): Piece[] => {
    if (seen.has(run)) return [];
    seen.add(run);
    const before = stats.runs;
    const pieces = processRun(run, styles, stats);
    if (stats.runs !== before) changed = true;
    return pieces;
  };

  for (const paragraph of collect(doc, "p")) {
    const pieces = runsInParagraph(paragraph).flatMap(consume);
    if (!capturePreview) continue;
    if (pieces.length === 0 || pieces.every((piece) => piece.raw === "" && piece.after === "")) {
      continue;
    }
    if (preview.length >= PREVIEW_PARAGRAPH_LIMIT) {
      truncated = true;
      continue;
    }
    preview.push({
      before: pieces.map((piece) => ({ text: piece.raw, ...piece.format })),
      after: pieces.map((piece) => piece.after).join(""),
    });
  }

  for (const run of collect(doc, "r")) consume(run);

  return {
    xml: changed ? serialize(doc, xml) : xml,
    changed,
    preview,
    truncated,
  };
}

function emptyStats(): ConvertStats {
  return {
    bold: 0,
    italic: 0,
    underline: 0,
    strike: 0,
    subscript: 0,
    superscript: 0,
    runs: 0,
  };
}

export async function convertDocx(input: ArrayBuffer | Uint8Array): Promise<ConvertResult> {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (bytes.byteLength === 0) throw new ConvertError(INVALID_MESSAGE, "invalid");

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    throw new ConvertError(INVALID_MESSAGE, "invalid");
  }

  const paths = Object.keys(zip.files).map((path) => path.replace(/^\/+/, ""));
  if (paths.includes("EncryptionInfo") || paths.includes("EncryptedPackage")) {
    throw new ConvertError(ENCRYPTED_MESSAGE, "encrypted");
  }
  if (!zip.file("word/document.xml")) {
    throw new ConvertError(INVALID_MESSAGE, "invalid");
  }

  const stylesFile = zip.file("word/styles.xml");
  const styles = loadStyles(stylesFile ? await stylesFile.async("string") : null);
  const stats = emptyStats();
  const preview: PreviewParagraph[] = [];
  let truncated = false;

  for (const path of paths) {
    if (!shouldProcessPart(path)) continue;
    const file = zip.file(path);
    if (!file || file.dir) continue;
    const xml = await file.async("string");
    const capturePreview = path === "word/document.xml";
    const result = transformPart(xml, styles, stats, capturePreview);
    if (result.changed) zip.file(path, result.xml);
    if (path === "word/document.xml") {
      preview.push(...result.preview);
      truncated = result.truncated;
    }
  }

  const data = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return { data, stats, preview, truncated };
}
