"use client";

import { useEffect, useId, useRef, useState, type ChangeEvent, type DragEvent, type ReactNode } from "react";
import type { ConvertResult, PreviewSpan } from "@/lib/convert-docx";

type Status = "idle" | "converting" | "done";

const SPECIMEN: { name: string; tag: string; sample: ReactNode }[] = [
  { name: "Bold", tag: "b", sample: <span className="font-semibold">ABR Loan</span> },
  { name: "Italic", tag: "i", sample: <span className="italic">per annum</span> },
  {
    name: "Underline",
    tag: "u",
    sample: <span className="underline decoration-ink/80 underline-offset-[3px]">Section 9.1</span>,
  },
  {
    name: "Strike",
    tag: "s",
    sample: <span className="line-through decoration-ink/70">deleted</span>,
  },
  {
    name: "Subscript",
    tag: "sub",
    sample: (
      <>
        H<sub className="text-[0.7em]">2</sub>O
      </>
    ),
  },
  {
    name: "Superscript",
    tag: "sup",
    sample: (
      <>
        n<sup className="text-[0.7em]">2</sup>
      </>
    ),
  },
];

const STAT_LABELS = [
  ["bold", "Bold"],
  ["italic", "Italic"],
  ["underline", "Underline"],
  ["strike", "Strike"],
  ["subscript", "Subscript"],
  ["superscript", "Superscript"],
] as const;

function taggedName(name: string): string {
  return `${name.replace(/\.docx$/i, "")}.tagged.docx`;
}

function formatBytes(size: number): string {
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${Math.max(1, Math.round(size / 1024))} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function yieldFrame(): Promise<void> {
  return new Promise((resolve) => {
    requestAnimationFrame(() => resolve());
  });
}

function downloadDocx(data: Uint8Array, filename: string) {
  const copy = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;
  const blob = new Blob([copy], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function failureMessage(error: unknown): string {
  if (error instanceof Error && error.name === "ConvertError") return error.message;
  return "The document couldn't be converted. It may be damaged.";
}

function spanClass(span: PreviewSpan): string {
  const classes: string[] = [];
  if (span.bold) classes.push("font-semibold");
  if (span.italic) classes.push("italic");
  if (span.underline) classes.push("underline decoration-ink/80 underline-offset-[3px]");
  if (span.strike) classes.push("line-through decoration-ink/70");
  if (span.subscript) classes.push("align-sub text-[0.75em]");
  if (span.superscript) classes.push("align-super text-[0.75em]");
  return classes.join(" ");
}

function TaggedText({ text }: { text: string }) {
  const tokens = text.split(/(<\/?(?:b|i|u|s|sub|sup)>)/g);
  return tokens.map((token, index) =>
    token.startsWith("<") ? (
      <span key={`${token}-${index}`} className="text-copper">
        {token}
      </span>
    ) : (
      <span key={`${index}-${token.slice(0, 12)}`}>{token}</span>
    ),
  );
}

function DocumentMark() {
  return (
    <svg width="36" height="44" viewBox="0 0 36 44" aria-hidden="true" className="text-ink">
      <path
        d="M6 2.5h16L30 10.5V41.5H6V2.5Z"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.25"
      />
      <path d="M22 2.5V10.5H30" fill="none" stroke="currentColor" strokeWidth="1.25" />
      <path d="M11 20h14M11 26h10M11 32h12" stroke="currentColor" strokeWidth="1.25" />
    </svg>
  );
}

export default function Converter() {
  const inputId = useId();
  const errorId = useId();
  const inputRef = useRef<HTMLInputElement>(null);
  const resultRef = useRef<HTMLElement>(null);
  const dragDepth = useRef(0);
  const generation = useRef(0);

  const [status, setStatus] = useState<Status>("idle");
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState("");
  const [fileName, setFileName] = useState("");
  const [fileSize, setFileSize] = useState(0);
  const [result, setResult] = useState<ConvertResult | null>(null);

  useEffect(() => {
    if (status !== "done" || !resultRef.current) return;
    const reduce = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    resultRef.current.scrollIntoView({
      behavior: reduce ? "auto" : "smooth",
      block: "start",
    });
  }, [status, result]);

  function reset() {
    generation.current += 1;
    dragDepth.current = 0;
    setDragging(false);
    setStatus("idle");
    setError(null);
    setNotice("");
    setResult(null);
    setFileName("");
    setFileSize(0);
    if (inputRef.current) inputRef.current.value = "";
    inputRef.current?.focus();
  }

  async function convertBytes(bytes: Uint8Array, name: string, size: number) {
    const id = ++generation.current;
    setStatus("converting");
    setError(null);
    setNotice("");
    setResult(null);
    setFileName(name);
    setFileSize(size);
    try {
      await yieldFrame();
      const { convertDocx } = await import("@/lib/convert-docx");
      const converted = await convertDocx(bytes);
      if (generation.current !== id) return;
      setResult(converted);
      setStatus("done");
      const runs = converted.stats.runs;
      setNotice(
        `${taggedName(name)} is ready. ${runs} ${runs === 1 ? "run" : "runs"} rewritten.`,
      );
    } catch (caught) {
      if (generation.current !== id) return;
      setStatus("idle");
      setResult(null);
      setError(failureMessage(caught));
    }
  }

  async function convertFile(file: File) {
    const name = file.name;
    const lower = name.toLowerCase();
    if (lower.endsWith(".doc") && !lower.endsWith(".docx")) {
      setError("Older .doc files need to be saved as .docx in Word before they can be converted.");
      setNotice("");
      return;
    }
    if (!lower.endsWith(".docx")) {
      setError("Only .docx files can be converted.");
      setNotice("");
      return;
    }
    const bytes = new Uint8Array(await file.arrayBuffer());
    await convertBytes(bytes, name, file.size);
  }

  async function onInputChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file || status === "converting") return;
    await convertFile(file);
  }

  function onDragEnter(event: DragEvent<HTMLDivElement>) {
    if (status === "converting") return;
    event.preventDefault();
    dragDepth.current += 1;
    setDragging(true);
  }

  function onDragOver(event: DragEvent<HTMLDivElement>) {
    if (status === "converting") return;
    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
  }

  function onDragLeave(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    dragDepth.current = Math.max(0, dragDepth.current - 1);
    if (dragDepth.current === 0) setDragging(false);
  }

  async function onDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    dragDepth.current = 0;
    setDragging(false);
    if (status === "converting") return;
    const file = event.dataTransfer.files?.[0];
    if (file) await convertFile(file);
  }

  async function trySample() {
    if (status === "converting") return;
    const id = ++generation.current;
    setStatus("converting");
    setError(null);
    setNotice("");
    setResult(null);
    setFileName("sample-clause.docx");
    setFileSize(0);
    try {
      await yieldFrame();
      const [{ convertDocx }, { buildSampleDocx }] = await Promise.all([
        import("@/lib/convert-docx"),
        import("@/lib/sample-docx"),
      ]);
      const sample = await buildSampleDocx();
      if (generation.current !== id) return;
      setFileSize(sample.byteLength);
      const converted = await convertDocx(sample);
      if (generation.current !== id) return;
      setResult(converted);
      setStatus("done");
      setNotice(
        `sample-clause.tagged.docx is ready. ${converted.stats.runs} runs rewritten.`,
      );
    } catch (caught) {
      if (generation.current !== id) return;
      setStatus("idle");
      setResult(null);
      setError(failureMessage(caught));
    }
  }

  const downloadName = fileName ? taggedName(fileName) : "document.tagged.docx";
  const busy = status === "converting";

  return (
    <div className="min-h-full text-ink">
      <p className="sr-only" aria-live="polite">
        {notice}
      </p>
      <header className="sticky top-0 z-20 border-b border-line/90 bg-paper/90 backdrop-blur-md">
        <div className="mx-auto flex max-w-6xl items-center justify-between gap-4 px-5 py-4 sm:px-8">
          <button
            type="button"
            onClick={reset}
            className="font-serif text-xl tracking-tight text-ink focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-4 focus-visible:outline-copper"
          >
            Word to HTML
          </button>
          <p className="flex items-center gap-2 text-xs tracking-wide text-muted">
            <span className="h-1.5 w-1.5 rounded-full bg-copper" aria-hidden="true" />
            Processed locally
          </p>
        </div>
      </header>

      <main className="mx-auto w-full max-w-6xl px-5 pb-24 pt-14 sm:px-8 sm:pt-20">
        <div className="max-w-3xl">
          <p className="text-xs font-medium uppercase tracking-[0.22em] text-copper">
            Word documents
          </p>
          <h1 className="mt-4 max-w-4xl font-serif text-5xl leading-[0.98] tracking-[-0.03em] text-ink sm:text-6xl lg:text-7xl">
            Formatting, written out as HTML.
          </h1>
          <p className="mt-6 max-w-xl text-lg leading-relaxed text-ink-soft">
            Drop a .docx. Bold, italic, underline, strike, subscript, and superscript come back as
            tags inside a new Word file.
          </p>
        </div>

        <div className="mt-14 grid items-start gap-10 lg:grid-cols-[minmax(0,1.35fr)_minmax(280px,0.8fr)] lg:gap-16">
          <div>
            <div
              onDragEnter={onDragEnter}
              onDragOver={onDragOver}
              onDragLeave={onDragLeave}
              onDrop={onDrop}
              className={`relative border bg-sheet px-6 py-14 shadow-[0_1px_0_rgba(26,23,20,0.04),0_30px_70px_-36px_rgba(26,23,20,0.45)] transition-colors sm:px-10 ${
                dragging ? "border-copper bg-[#f8f1e7]" : "border-line"
              }`}
            >
              <div className="pointer-events-none absolute inset-3 border border-line" />
              <label
                htmlFor={inputId}
                className={`relative flex min-h-52 cursor-pointer flex-col items-center justify-center text-center focus-within:outline focus-within:outline-2 focus-within:outline-offset-8 focus-within:outline-copper ${
                  busy ? "cursor-wait" : ""
                }`}
              >
                <input
                  ref={inputRef}
                  id={inputId}
                  type="file"
                  accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                  className="sr-only"
                  disabled={busy}
                  aria-describedby={error ? errorId : undefined}
                  aria-busy={busy}
                  onChange={onInputChange}
                />
                <DocumentMark />
                {busy ? (
                  <div className="mt-6 w-full max-w-sm">
                    <p className="font-serif text-3xl tracking-tight">Reading the document</p>
                    <p className="mt-2 truncate text-sm text-muted">
                      {fileName || "Working through each run"}
                    </p>
                    <div className="mt-6 h-px overflow-hidden bg-line">
                      <div className="progress-bar h-full w-1/3 bg-copper" />
                    </div>
                  </div>
                ) : (
                  <div className="mt-6">
                    <p className="font-serif text-3xl tracking-tight">Drop a .docx here</p>
                    <p className="mt-2 text-sm text-muted">or browse from this computer</p>
                  </div>
                )}
              </label>
            </div>

            <div className="mt-4 flex flex-wrap items-center justify-between gap-3">
              <button
                type="button"
                onClick={trySample}
                disabled={busy}
                className="text-sm text-ink underline decoration-line underline-offset-4 hover:decoration-copper focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-4 focus-visible:outline-copper disabled:cursor-wait disabled:opacity-60"
              >
                Try the sample clause
              </button>
              <p className="text-xs tracking-wide text-muted">Nothing is uploaded</p>
            </div>

            {error ? (
              <p
                id={errorId}
                role="alert"
                className="mt-4 border border-[#e4cbbd] bg-[#fbf6f2] px-4 py-3 text-sm leading-relaxed text-copper-deep"
              >
                {error}
              </p>
            ) : null}
          </div>

          <aside className="lg:pt-2">
            <h2 className="font-serif text-2xl tracking-tight">The six marks</h2>
            <p className="mt-2 text-sm leading-relaxed text-muted">
              Each formatted run is written out in place. The rest of the document — tables,
              headers, footnotes — stays as it was.
            </p>
            <ul className="mt-6">
              {SPECIMEN.map((item) => (
                <li
                  key={item.tag}
                  className="grid grid-cols-[6.25rem_minmax(0,1fr)_auto] items-baseline gap-3 border-t border-line py-3"
                >
                  <span className="text-[11px] uppercase tracking-[0.16em] text-muted">
                    {item.name}
                  </span>
                  <span className="truncate text-[15px]">{item.sample}</span>
                  <code className="font-mono text-[13px] text-copper">&lt;{item.tag}&gt;</code>
                </li>
              ))}
              <li className="border-t border-line" />
            </ul>

            <figure className="mt-8 border border-line bg-sheet px-5 py-5">
              <figcaption className="text-[11px] uppercase tracking-[0.16em] text-muted">
                From the clause
              </figcaption>
              <p className="mt-3 font-serif text-[1.35rem] leading-snug">
                each <span className="font-sans text-[1.05rem] font-semibold">ABR Loan</span> that
                is an{" "}
                <span className="font-sans text-[1.05rem] underline decoration-ink/80 underline-offset-[3px]">
                  Initial
                </span>
              </p>
              <p className="mt-3 font-mono text-[13px] leading-relaxed text-ink-soft">
                each <span className="text-copper">&lt;b&gt;</span>ABR Loan
                <span className="text-copper">&lt;/b&gt;</span> that is an{" "}
                <span className="text-copper">&lt;u&gt;</span>Initial
                <span className="text-copper">&lt;/u&gt;</span>
              </p>
            </figure>
            <p className="mt-4 text-sm leading-relaxed text-muted">
              Adjacent runs stay separate, so two underlined words can become two tags. A space
              that lives inside a run stays inside the tag.
            </p>
          </aside>
        </div>

        {result ? (
          <section ref={resultRef} className="mt-20 scroll-mt-24 border-t border-line pt-12">
            <div className="flex flex-col gap-8 sm:flex-row sm:items-end sm:justify-between">
              <div className="min-w-0">
                <p className="text-[11px] uppercase tracking-[0.18em] text-copper">Ready</p>
                <h2 className="mt-2 break-words font-serif text-4xl tracking-tight sm:text-5xl">
                  {downloadName}
                </h2>
                <p className="mt-3 text-sm text-muted">
                  {fileName}
                  {fileSize > 0 ? ` · ${formatBytes(fileSize)}` : ""} · {result.stats.runs}{" "}
                  {result.stats.runs === 1 ? "run" : "runs"} rewritten
                </p>
              </div>
              <div className="flex flex-wrap items-center gap-4">
                <button
                  type="button"
                  onClick={() => downloadDocx(result.data, downloadName)}
                  className="inline-flex h-12 items-center bg-ink px-6 text-sm font-medium text-sheet transition-colors hover:bg-copper-deep focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-copper"
                >
                  Download .docx
                </button>
                <button
                  type="button"
                  onClick={reset}
                  className="h-12 text-sm text-muted underline decoration-line underline-offset-4 hover:text-ink focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-4 focus-visible:outline-copper"
                >
                  Convert another
                </button>
              </div>
            </div>

            {result.stats.runs === 0 ? (
              <p className="mt-8 max-w-2xl text-sm leading-relaxed text-ink-soft">
                No bold, italic, underline, strike, subscript, or superscript was found. The
                download matches the original wording.
              </p>
            ) : null}

            <dl className="mt-8 grid grid-cols-2 gap-px overflow-hidden border border-line bg-line sm:grid-cols-3 lg:grid-cols-6">
              {STAT_LABELS.map(([key, label]) => (
                <div key={key} className="bg-sheet px-4 py-4">
                  <dt className="text-[11px] uppercase tracking-[0.16em] text-muted">{label}</dt>
                  <dd className="mt-2 font-serif text-3xl tabular-nums tracking-tight">
                    {result.stats[key]}
                  </dd>
                </div>
              ))}
            </dl>

            <div className="mt-8 grid gap-6 lg:grid-cols-2">
              <article className="border border-line bg-sheet">
                <header className="border-b border-line px-5 py-3">
                  <h3 className="text-[11px] uppercase tracking-[0.16em] text-muted">Original</h3>
                </header>
                <div className="max-h-[32rem] overflow-auto px-5 py-5">
                  {result.preview.length === 0 ? (
                    <p className="text-sm leading-relaxed text-muted">
                      No body text to preview. Headers, footnotes, and comments are still converted
                      in the download.
                    </p>
                  ) : (
                    result.preview.map((paragraph, index) => (
                      <p
                        key={`before-${index}`}
                        className="mb-4 whitespace-pre-wrap text-[15px] leading-relaxed last:mb-0"
                      >
                        {paragraph.before.map((span, spanIndex) => (
                          <span key={`span-${spanIndex}`} className={spanClass(span)}>
                            {span.text}
                          </span>
                        ))}
                      </p>
                    ))
                  )}
                </div>
              </article>

              <article className="border border-line bg-sheet">
                <header className="border-b border-line px-5 py-3">
                  <h3 className="text-[11px] uppercase tracking-[0.16em] text-muted">With tags</h3>
                </header>
                <div className="max-h-[32rem] overflow-auto px-5 py-5">
                  {result.preview.length === 0 ? (
                    <p className="text-sm leading-relaxed text-muted">
                      The tagged file is ready even when the body has nothing to show here.
                    </p>
                  ) : (
                    result.preview.map((paragraph, index) => (
                      <p
                        key={`after-${index}`}
                        className="mb-4 whitespace-pre-wrap font-mono text-[13px] leading-relaxed text-ink last:mb-0"
                      >
                        <TaggedText text={paragraph.after} />
                      </p>
                    ))
                  )}
                </div>
              </article>
            </div>

            {result.truncated ? (
              <p className="mt-4 text-sm text-muted">
                Showing the first 40 paragraphs. The downloaded file includes the full document.
              </p>
            ) : null}
          </section>
        ) : null}
      </main>
    </div>
  );
}
