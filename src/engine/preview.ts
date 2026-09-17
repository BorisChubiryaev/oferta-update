// Предпросмотр правок.
//
// Раньше предпросмотр угадывал место правки собственным поиском, а движок
// применял её по своей логике. Расхождение было неизбежным и однажды дало
// худший из возможных исходов: правка попала в итоговый файл, но на шаге
// «Проверка» её не было видно. Поэтому предпросмотр теперь ПРОГОНЯЕТ ровно тот
// же движок в том же порядке и показывает абзац, который реально получился.
import { applyOneOp, applicationOrder, type ApplyState } from "./apply";
import { decodeXml } from "./ooxml";
import { indexFootnotes, findFootnoteById } from "./offer-index";
import type { BuildOptions, Operation } from "./types";

export type SegmentMark = "keep" | "ins" | "del";

export interface PreviewSegment {
  text: string;
  mark: SegmentMark;
}

export interface PreviewSnippet {
  /** Применится ли правка. */
  ok: boolean;
  kind: "paragraph" | "table" | "manual" | "none";
  /** Сообщение движка — ровно то, что попадёт в отчёт о сборке. */
  message: string;
  /** Итоговый абзац, разложенный на «оставлено / вставлено / удалено». */
  segments?: PreviewSegment[];
}

const CTX = 140; // сколько символов неизменного текста показывать вокруг правки

const RUN_RE = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;
const WT_RE = /<w:(?:t|delText)(?:\s[^>]*)?>([\s\S]*?)<\/w:(?:t|delText)>/g;

/** Границы абзаца, внутри которого лежит позиция pos. */
function enclosingParagraph(xml: string, pos: number): { start: number; end: number } | null {
  const openTag = xml.lastIndexOf("<w:p ", pos);
  const openPlain = xml.lastIndexOf("<w:p>", pos);
  const start = Math.max(openTag, openPlain);
  if (start < 0) return null;
  const end = xml.indexOf("</w:p>", start);
  if (end < 0) return null;
  return { start, end: end + "</w:p>".length };
}

function runText(runXml: string): string {
  const parts: string[] = [];
  let m: RegExpExecArray | null;
  WT_RE.lastIndex = 0;
  while ((m = WT_RE.exec(runXml)) !== null) parts.push(decodeXml(m[1]));
  return parts.join("");
}

/** Разложить абзац на отрезки «оставлено / вставлено / удалено». */
function segmentsOf(paragraphXml: string, opts: BuildOptions): PreviewSegment[] {
  const color = opts.highlightColor ?? "EE0000";
  // <w:ins>/<w:del> оборачивают раны — заранее помечаем занятые ими диапазоны.
  const wrappers = new Map<number, SegmentMark>();
  for (const [tag, mark] of [
    ["ins", "ins"],
    ["del", "del"],
  ] as const) {
    const re = new RegExp(`<w:${tag}\\b[^>]*>[\\s\\S]*?</w:${tag}>`, "g");
    let w: RegExpExecArray | null;
    while ((w = re.exec(paragraphXml)) !== null) {
      for (let i = w.index; i < w.index + w[0].length; i++) wrappers.set(i, mark);
    }
  }
  const out: PreviewSegment[] = [];
  let m: RegExpExecArray | null;
  RUN_RE.lastIndex = 0;
  while ((m = RUN_RE.exec(paragraphXml)) !== null) {
    const text = runText(m[0]);
    if (!text) continue;
    const wrapper = wrappers.get(m.index);
    let mark: SegmentMark = "keep";
    if (wrapper === "del" || /<w:strike\s*\/>/.test(m[0]) || /<w:delText/.test(m[0])) mark = "del";
    else if (wrapper === "ins" || m[0].includes(`w:color w:val="${color}"`)) mark = "ins";
    const last = out[out.length - 1];
    if (last && last.mark === mark) last.text += text;
    else out.push({ text, mark });
  }
  return out;
}

/** Обрезать неизменный текст вокруг правки, чтобы карточка оставалась читаемой. */
function trimContext(segments: PreviewSegment[]): PreviewSegment[] {
  const first = segments.findIndex((s) => s.mark !== "keep");
  if (first < 0) {
    const all = segments.map((s) => s.text).join("");
    return [{ mark: "keep", text: all.length > CTX * 2 ? all.slice(0, CTX * 2) + "…" : all }];
  }
  const reversed = [...segments].reverse().findIndex((s) => s.mark !== "keep");
  const last = segments.length - 1 - reversed;
  return segments.map((s, i) => {
    if (s.mark !== "keep") return s;
    if (i < first) return { mark: s.mark, text: s.text.length > CTX ? "…" + s.text.slice(-CTX) : s.text };
    if (i > last) return { mark: s.mark, text: s.text.length > CTX ? s.text.slice(0, CTX) + "…" : s.text };
    return s;
  });
}

const TABLE_OPS = new Set([
  "append_table_rows",
  "replace_table_rows",
  "sort_table_alpha",
  "insert_table_row_alpha",
]);

/**
 * Прогнать все операции и вернуть предпросмотр по каждой.
 *
 * Прогон именно накопительный: правки одного пакета ссылаются на нумерацию,
 * которая получается после предыдущих, и показывать каждую «по отдельности на
 * исходном документе» значило бы показывать неправду.
 */
export function previewOperations(
  offer: Omit<ApplyState, "document"> & { document: string },
  operations: Operation[],
  opts: BuildOptions = {},
): Map<string, PreviewSnippet> {
  const state: ApplyState = {
    document: offer.document,
    footnotes: offer.footnotes,
    numbering: offer.numbering,
    styles: offer.styles,
  };
  const result = new Map<string, PreviewSnippet>();
  for (const i of applicationOrder(operations)) {
    const op = operations[i];
    const res = applyOneOp(op, state, opts);

    if (op.type === "manual") {
      result.set(op.id, { ok: false, kind: "manual", message: op.note ?? res.message });
      continue;
    }
    if (!res.ok) {
      result.set(op.id, { ok: false, kind: "none", message: res.message });
      continue;
    }
    if (TABLE_OPS.has(op.type)) {
      result.set(op.id, { ok: true, kind: "table", message: res.message });
      continue;
    }
    // Сноски живут в footnotes.xml, а не в document.xml — обычный поиск
    // «абзац вокруг orderKey» тут не сработает. Ищем нужную сноску отдельно,
    // чтобы оператор видел «было / стало» и для правок сносок, а не только
    // сообщение движка.
    if (op.target.kind === "footnote" && state.footnotes) {
      const idxAfter = indexFootnotes(state.document);
      // Номер сноски: обычно он есть прямо в target; для варианта «сноска
      // к пункту» (atPoint) реальный номер разрешается внутри applyOneOp и
      // виден здесь только через orderKey (позицию ссылки в теле) —
      // сопоставляем по ней.
      const fnNumber =
        op.target.number ||
        [...idxAfter.displayToBodyPos.entries()].find(([, pos]) => pos === res.orderKey)?.[0];
      const id = fnNumber ? idxAfter.displayToId.get(fnNumber) : undefined;
      const fn = id !== undefined ? findFootnoteById(state.footnotes, id) : null;
      if (fn) {
        result.set(op.id, {
          ok: true,
          kind: "paragraph",
          message: res.message,
          segments: trimContext(segmentsOf(fn.inner, opts)),
        });
        continue;
      }
      result.set(op.id, { ok: true, kind: "table", message: res.message });
      continue;
    }
    const span = enclosingParagraph(state.document, Math.max(0, res.orderKey));
    if (!span) {
      result.set(op.id, { ok: true, kind: "table", message: res.message });
      continue;
    }
    result.set(op.id, {
      ok: true,
      kind: "paragraph",
      message: res.message,
      segments: trimContext(segmentsOf(state.document.slice(span.start, span.end), opts)),
    });
  }
  return result;
}
