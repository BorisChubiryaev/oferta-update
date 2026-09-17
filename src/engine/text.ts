// Текстовые утилиты для разбора инструкций.
import { decodeXml } from "./ooxml";

const WT_RE = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
const P_RE = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;
const TBL_RE = /<w:tbl>[\s\S]*?<\/w:tbl>/g;

/** Видимый текст одного абзаца/фрагмента XML. */
export function xmlText(fragment: string): string {
  const parts: string[] = [];
  let m: RegExpExecArray | null;
  WT_RE.lastIndex = 0;
  while ((m = WT_RE.exec(fragment)) !== null) parts.push(decodeXml(m[1]));
  return parts.join("").replace(/\u00A0/g, " ").trim();
}

/** Все абзацы документа как массив строк (пустые отфильтрованы). */
export function paragraphs(documentXml: string): string[] {
  const out: string[] = [];
  let m: RegExpExecArray | null;
  P_RE.lastIndex = 0;
  while ((m = P_RE.exec(documentXml)) !== null) {
    const t = xmlText(m[0]);
    if (t) out.push(t);
  }
  return out;
}

/**
 * Абзацы ВНЕ таблиц. Ячейка таблицы состоит из тех же <w:p>, поэтому обычный
 * разбор смешивает инструкции с содержимым таблиц: строка «ООО «СберЛогистика»»
 * попадает в поток инструкций и мешает склейке многоабзацных редакций.
 */
export function paragraphsOutsideTables(documentXml: string): string[] {
  const gaps: { from: number; to: number }[] = [];
  let t: RegExpExecArray | null;
  TBL_RE.lastIndex = 0;
  while ((t = TBL_RE.exec(documentXml)) !== null) {
    gaps.push({ from: t.index, to: t.index + t[0].length });
  }
  const inTable = (pos: number) => gaps.some((g) => pos >= g.from && pos < g.to);
  const out: string[] = [];
  let m: RegExpExecArray | null;
  P_RE.lastIndex = 0;
  while ((m = P_RE.exec(documentXml)) !== null) {
    if (inTable(m.index)) continue;
    const txt = xmlText(m[0]);
    if (txt) out.push(txt);
  }
  return out;
}

/**
 * Привести все виды кавычек к «ёлочкам».
 *
 * Разбор инструкций опирается на «…» — в них лежит новая редакция. Но часть
 * документов «Изменения» набрана лапками ("Публичная Оферта", "9.6 Клиент
 * проинформирован…"), и для парсера такой текст был невидим: правка
 * распознавалась как «не найден текст новой редакции».
 *
 * Длина строки сохраняется (замена всегда символ-в-символ), потому что разбор
 * работает с позициями символов в этой же строке.
 *
 * Прямая кавычка не различает открывающую и закрывающую, поэтому направление
 * определяется соседом слева — как это делают «умные кавычки» в текстовых
 * редакторах: после пробела или открывающей скобки кавычка открывающая, после
 * буквы или знака препинания — закрывающая. Считать их просто по очереди
 * нельзя: в документах сплошь и рядом встречается вложенность («Оферта
 * "Удобный доступ"») и пропущенная закрывающая кавычка, и одна лишняя кавычка
 * переворачивала бы все последующие.
 */
export function normalizeQuotes(s: string): string {
  const out = Array.from(s.replace(/[“„‟]/g, "«").replace(/[”]/g, "»"));
  for (let i = 0; i < out.length; i++) {
    if (out[i] !== '"') continue;
    const prev = i > 0 ? out[i - 1] : "";
    out[i] = prev === "" || /[\s([{«\-–—/]/.test(prev) ? "«" : "»";
  }
  return out.join("");
}

/**
 * Извлечь содержимое сбалансированных кавычек-«ёлочек», начиная с позиции
 * первого «. Учитывает вложенность («ООО СК «Сбербанк Страхование»»).
 */
export function extractGuillemet(
  s: string,
  fromIndex: number,
): { content: string; endIndex: number; balanced: boolean } | null {
  const open = s.indexOf("«", fromIndex);
  if (open < 0) return null;
  let depth = 0;
  for (let i = open; i < s.length; i++) {
    if (s[i] === "«") depth++;
    else if (s[i] === "»") {
      depth--;
      if (depth === 0) {
        return { content: s.slice(open + 1, i), endIndex: i, balanced: true };
      }
    }
  }
  // Не закрыто — берём до конца, но честно сообщаем об этом: по одному лишь
  // endIndex случай «кавычка не закрыта» неотличим от «кавычка закрывается
  // последним символом», а решения у них противоположные.
  return { content: s.slice(open + 1), endIndex: s.length - 1, balanced: false };
}

function tableRows(tableXml: string): string[][] {
  const rows: string[][] = [];
  const trRe = /<w:tr\b[\s\S]*?<\/w:tr>/g;
  let tr: RegExpExecArray | null;
  while ((tr = trRe.exec(tableXml)) !== null) {
    const cells: string[] = [];
    const tcRe = /<w:tc>[\s\S]*?<\/w:tc>/g;
    let tc: RegExpExecArray | null;
    while ((tc = tcRe.exec(tr[0])) !== null) {
      // Ячейка нередко состоит из нескольких абзацев — наименование компании
      // и её сайт. Склеив их подряд, получаем «ООО «Ромашка»https://…»,
      // поэтому границы абзацев сохраняем переводом строки.
      const parts = (tc[0].match(P_RE) ?? []).map((x) => xmlText(x)).filter(Boolean);
      cells.push(parts.length ? parts.join("\n") : xmlText(tc[0]));
    }
    rows.push(cells);
  }
  return rows;
}

/** Все таблицы документа как массив строк-ячеек. */
export function tables(documentXml: string): string[][][] {
  const result: string[][][] = [];
  let t: RegExpExecArray | null;
  TBL_RE.lastIndex = 0;
  while ((t = TBL_RE.exec(documentXml)) !== null) result.push(tableRows(t[0]));
  return result;
}

/** Абзац или таблица — в том порядке, в каком они идут в документе. */
export type DocBlock =
  | { kind: "p"; text: string }
  | { kind: "tbl"; rows: string[][] };

/**
 * Документ единым потоком блоков.
 *
 * Отдельно список абзацев и отдельно список таблиц теряют главное — КТО К ЧЕМУ
 * относится. А в документах «Изменения» содержимое правки сплошь и рядом лежит
 * не в кавычках, а в таблице сразу под инструкцией («…добавить следующие
 * строки:» + таблица). Без порядка следования приходилось угадывать таблицу по
 * её содержимому, и при двух таблицах в одном документе выбор был случайным.
 */
export function documentBlocks(documentXml: string): DocBlock[] {
  const spans: { from: number; to: number; rows: string[][] }[] = [];
  let t: RegExpExecArray | null;
  TBL_RE.lastIndex = 0;
  while ((t = TBL_RE.exec(documentXml)) !== null) {
    spans.push({ from: t.index, to: t.index + t[0].length, rows: tableRows(t[0]) });
  }
  const blocks: { at: number; block: DocBlock }[] = spans.map((s) => ({
    at: s.from,
    block: { kind: "tbl", rows: s.rows } as DocBlock,
  }));
  let m: RegExpExecArray | null;
  P_RE.lastIndex = 0;
  while ((m = P_RE.exec(documentXml)) !== null) {
    if (spans.some((s) => m!.index >= s.from && m!.index < s.to)) continue;
    const text = xmlText(m[0]);
    if (text) blocks.push({ at: m.index, block: { kind: "p", text } });
  }
  return blocks.sort((a, b) => a.at - b.at).map((b) => b.block);
}
