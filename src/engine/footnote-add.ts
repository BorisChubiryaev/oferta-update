// Добавление НОВОЙ сноски. Сноски Word нумеруются по порядку ссылок
// <w:footnoteReference> в теле, поэтому достаточно вставить новую ссылку в
// нужном месте и добавить элемент <w:footnote> — последующие сноски
// перенумеровываются автоматически.
import { renderInsertRuns } from "./render";
import type { BuildOptions } from "./types";

/** Максимальный внутренний id сноски (для генерации нового уникального id). */
export function maxFootnoteId(footnotesXml: string): number {
  const ids = [...footnotesXml.matchAll(/<w:footnote\b[^>]*\bw:id="(-?\d+)"/g)].map((m) =>
    parseInt(m[1], 10),
  );
  return ids.length ? Math.max(...ids) : 1;
}

const DEFAULT_REF_RPR = '<w:rPr><w:rStyle w:val="af4"/><w:vertAlign w:val="superscript"/></w:rPr>';

/**
 * Оформление рана-ссылки на сноску — берём у первой такой ссылки в документе.
 *
 * Ран находим ЦЕЛИКОМ и только потом достаём из него <w:rPr>. Попытка сделать
 * это одним выражением приводила к катастрофе: ленивый квантификатор с
 * возвратами захватывал кусок документа в несколько абзацев, и этот кусок
 * вставлялся в текст как «свойства рана» — Оферта получала дубли заголовков и
 * лишние разделы. Отсюда же и проверка формы результата ниже: подставлять в
 * документ произвольную строку нельзя.
 */
export function footnoteRefRunRpr(documentXml: string): string {
  const runRe = /<w:r\b[^>]*>[\s\S]*?<\/w:r>/g;
  let m: RegExpExecArray | null;
  while ((m = runRe.exec(documentXml)) !== null) {
    if (!m[0].includes("<w:footnoteReference")) continue;
    const rPr = m[0].match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
    if (rPr && rPr[0].length < 400) return rPr[0];
    return DEFAULT_REF_RPR;
  }
  return DEFAULT_REF_RPR;
}

/** Ран-ссылка на новую сноску для вставки в тело. */
export function buildFootnoteReferenceRun(id: number, rPr: string): string {
  return `<w:r>${rPr}<w:footnoteReference w:id="${id}"/></w:r>`;
}

/** Элемент <w:footnote> с выделенным текстом. */
export function buildFootnoteElement(id: number, text: string, opts: BuildOptions): string {
  const body = renderInsertRuns(" " + text.trim(), opts);
  return (
    `<w:footnote w:id="${id}">` +
    `<w:p><w:pPr><w:pStyle w:val="af2"/><w:jc w:val="both"/></w:pPr>` +
    `<w:r><w:rPr><w:rStyle w:val="af4"/></w:rPr><w:footnoteRef/></w:r>` +
    `${body}</w:p></w:footnote>`
  );
}

/** Добавить элемент сноски перед закрывающим тегом. */
export function appendFootnoteElement(footnotesXml: string, element: string): string {
  return footnotesXml.replace(/<\/w:footnotes>\s*$/, element + "</w:footnotes>");
}
