// Индексация Оферты: сопоставление «видимых» номеров сносок с их id,
// а также границы отдельных сносок внутри footnotes.xml.
//
// Сноски в Word нумеруются автоматически — по порядку ссылок
// <w:footnoteReference> в теле документа. «Сноска № 32» из инструкции — это
// 32-я по счёту ссылка в тексте, а не элемент с w:id="32".

export interface FootnoteIndex {
  /** видимый номер (1..N) -> внутренний w:id */
  displayToId: Map<number, number>;
  /** порядковая позиция ссылки в теле (для сортировки) */
  displayToBodyPos: Map<number, number>;
}

export function indexFootnotes(documentXml: string): FootnoteIndex {
  const displayToId = new Map<number, number>();
  const displayToBodyPos = new Map<number, number>();
  const re = /<w:footnoteReference[^>]*\bw:id="(\d+)"/g;
  let m: RegExpExecArray | null;
  let display = 0;
  while ((m = re.exec(documentXml)) !== null) {
    display += 1; // видимая нумерация начинается с 1
    displayToId.set(display, parseInt(m[1], 10));
    displayToBodyPos.set(display, m.index);
  }
  return { displayToId, displayToBodyPos };
}

export interface FootnoteBlock {
  id: number;
  start: number;
  end: number;
  inner: string;
}

/** Все сноски документа (для поиска по содержимому). */
export function allFootnotes(footnotesXml: string): FootnoteBlock[] {
  const re = /<w:footnote\b[^>]*\bw:id="(-?\d+)"[^>]*>([\s\S]*?)<\/w:footnote>/g;
  const out: FootnoteBlock[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(footnotesXml)) !== null) {
    out.push({ id: parseInt(m[1], 10), start: m.index, end: m.index + m[0].length, inner: m[2] });
  }
  return out;
}

/** Границы конкретной сноски (по её w:id) внутри footnotes.xml. */
export function findFootnoteById(
  footnotesXml: string,
  id: number,
): { start: number; end: number; inner: string } | null {
  const re = new RegExp(`<w:footnote\\b[^>]*\\bw:id="${id}"[^>]*>([\\s\\S]*?)<\\/w:footnote>`);
  const m = footnotesXml.match(re);
  if (!m || m.index === undefined) return null;
  return { start: m.index, end: m.index + m[0].length, inner: m[1] };
}
