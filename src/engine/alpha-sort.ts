// Алфавитная пересортировка таблицы Приложения и вставка строки «по алфавиту».
//
// Ключ сортировки выведен эмпирически из действующего Приложения 1 и точно
// воспроизводит порядок, выстроенный юристами (49/49 совпадений):
//   • отбрасываем организационную форму (ООО/АО/АНО…) в конце ИЛИ в начале —
//     в этом приложении форма пишется после названия: «БИЗОН» ООО;
//   • убираем кавычки («» и "), регистр не важен, ё → е;
//   • пунктуацию (дефис) СОХРАНЯЕМ: «С-МАРКЕТИНГ» идёт перед «СалютДевайсы».
// Префиксы вроде НКО/ПКО/СК частью названия и остаются в ключе.
import { escapeXml, tableCellText } from "./ooxml";

const WT_RE = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;

/** Организационно-правовые формы, отбрасываемые из ключа сортировки. */
export const LEGAL_FORMS = [
  "ООО", "ОАО", "ЗАО", "ПАО", "АО", "АНО", "НАО", "ИП", "ПК", "ГК",
];

const FORMS_RE = LEGAL_FORMS.join("|");

/** Ключ сортировки названия компании. */
export function sortKey(name: string): string {
  let t = name.trim();
  t = t.replace(new RegExp(`\\s*(?:${FORMS_RE})\\s*$`), "");
  t = t.replace(new RegExp(`^(?:${FORMS_RE})\\s+`), "");
  t = t.replace(/["«»“”„]/g, "");
  // Внутренние пробелы схлопываем: в документах «Изменения» встречаются
  // двойные пробелы («Объединенное  Кредитное Бюро»), из-за которых
  // наименование иначе не совпало бы с текстом Оферты.
  return t.replace(/\s+/g, " ").trim().toLowerCase().replace(/ё/g, "е");
}

export interface TableRow {
  xml: string;
  cells: string[]; // видимый текст ячеек
}

const cellText = tableCellText;

/** Разобрать таблицу на строки с текстом ячеек. */
export function parseRows(tableInner: string): TableRow[] {
  const trs = tableInner.match(/<w:tr\b[\s\S]*?<\/w:tr>/g) ?? [];
  return trs.map((xml) => ({
    xml,
    cells: (xml.match(/<w:tc>[\s\S]*?<\/w:tc>/g) ?? []).map(cellText),
  }));
}

/**
 * Настоящее вертикальное объединение — это ячейка-продолжение: <w:vMerge/>
 * без w:val либо w:val="continue". Одиночный "restart" без продолжения —
 * безобидный артефакт редактирования, перестановке строк он не мешает.
 */
export function hasRealVMerge(tableInner: string): boolean {
  const tags = tableInner.match(/<w:vMerge\b[^>]*\/?>/g) ?? [];
  return tags.some((t) => !/w:val="restart"/.test(t));
}

/** Заменить номер в первой ячейке строки, сохранив всё форматирование. */
export function setRowNumber(trXml: string, num: number): string {
  const tcMatch = trXml.match(/<w:tc>[\s\S]*?<\/w:tc>/);
  if (!tcMatch) return trXml;
  let first = true;
  const newTc = tcMatch[0].replace(WT_RE, (full, _inner, offset) => {
    void _inner;
    void offset;
    const attrs = full.match(/^<w:t(\s[^>]*)?>/)?.[1] ?? "";
    if (first) {
      first = false;
      return `<w:t${attrs}>${escapeXml(String(num))}</w:t>`;
    }
    return `<w:t${attrs}></w:t>`;
  });
  return trXml.replace(tcMatch[0], newTc);
}

export interface SortResult {
  xml: string;
  moves: { name: string; from: number; to: number }[];
  order: { number: number; name: string }[];
  warnings: string[];
}

/**
 * Отсортировать строки таблицы по алфавиту и перенумеровать их.
 * `nameCol` — индекс колонки с наименованием (по умолчанию 1).
 */
export function sortTableAlphabetically(
  tableInner: string,
  nameCol = 1,
): SortResult | { error: string } {
  if (hasRealVMerge(tableInner)) {
    return { error: "в таблице есть объединённые по вертикали ячейки — сортировка выполняется вручную" };
  }
  const rows = parseRows(tableInner);
  if (rows.length < 2) return { error: "в таблице недостаточно строк для сортировки" };

  // Шапка — строка, у которой первая ячейка не является числом.
  const hasHeader = !/^\d+$/.test(rows[0].cells[0] ?? "");
  const header = hasHeader ? rows.slice(0, 1) : [];
  const body = hasHeader ? rows.slice(1) : rows;

  const warnings: string[] = [];
  const withKeys = body.map((r, i) => ({
    row: r,
    origIndex: i + 1,
    name: r.cells[nameCol] ?? "",
    key: sortKey(r.cells[nameCol] ?? ""),
  }));
  const empty = withKeys.filter((w) => !w.key).length;
  if (empty) warnings.push(`строк без наименования: ${empty} — проверьте порядок вручную`);

  const sorted = [...withKeys].sort((a, b) =>
    a.key < b.key ? -1 : a.key > b.key ? 1 : a.origIndex - b.origIndex,
  );

  const moves = sorted
    .map((w, i) => ({ name: w.name, from: w.origIndex, to: i + 1 }))
    .filter((m) => m.from !== m.to);

  const bodyXml = sorted.map((w, i) => setRowNumber(w.row.xml, i + 1)).join("");
  const headerXml = header.map((h) => h.xml).join("");
  // Пересобираем таблицу: свойства таблицы (до первой строки) + строки.
  const firstTr = tableInner.indexOf("<w:tr");
  const prefix = firstTr >= 0 ? tableInner.slice(0, firstTr) : tableInner;
  const suffix = tableInner.endsWith("</w:tbl>") ? "</w:tbl>" : "";
  const xml = prefix + headerXml + bodyXml + suffix;

  return {
    xml,
    moves,
    order: sorted.map((w, i) => ({ number: i + 1, name: w.name })),
    warnings,
  };
}

/**
 * Шапка ли первая строка.
 *
 * `known` задаёт ответ явно и нужен вызывающему, который шапку уже отрезал.
 * Автоопределение «первая ячейка — не число» на таком куске ошибается: у ТОЛЬКО
 * ЧТО вставленной строки номер ещё не проставлен (его присваивают в конце, при
 * сквозной перенумерации), и она принималась за шапку — то есть выпадала из
 * сравнения. Из-за этого вторая вставляемая компания не видела первую и
 * вставала ПЕРЕД ней: «Лента.Ру» оказывалась выше «Газета.Ру».
 */
function headerRows(rows: TableRow[], known?: boolean): number {
  const has = known ?? !/^\d+$/.test(rows[0]?.cells[0] ?? "");
  return has ? 1 : 0;
}

/** Есть ли уже строка с таким наименованием (по ключу сортировки)? */
export function findExistingRow(
  tableInner: string,
  name: string,
  nameCol = 1,
  hasHeader?: boolean,
): { number: string; name: string } | null {
  const rows = parseRows(tableInner);
  const body = rows.slice(headerRows(rows, hasHeader));
  const key = sortKey(name);
  if (!key) return null;
  const hit = body.find((r) => sortKey(r.cells[nameCol] ?? "") === key);
  return hit ? { number: hit.cells[0] ?? "", name: hit.cells[nameCol] ?? "" } : null;
}

/** Найти позицию (1-based) для вставки нового наименования по алфавиту. */
export function alphabeticalPosition(
  tableInner: string,
  name: string,
  nameCol = 1,
  hasHeader?: boolean,
): number {
  const rows = parseRows(tableInner);
  const body = rows.slice(headerRows(rows, hasHeader));
  const key = sortKey(name);
  let pos = body.length + 1;
  for (let i = 0; i < body.length; i++) {
    if (sortKey(body[i].cells[nameCol] ?? "") > key) {
      pos = i + 1;
      break;
    }
  }
  return pos;
}

/**
 * Упорядочена ли таблица по алфавиту.
 *
 * Вставка «по алфавиту» в неупорядоченную таблицу даёт формально верное, но
 * бессмысленное место: строка встаёт перед первым наименованием, которое больше
 * неё, а дальше порядок всё равно произвольный. Оператор должен об этом знать.
 */
export function isAlphabeticallyOrdered(
  tableInner: string,
  nameCol = 1,
  hasHeader?: boolean,
): boolean {
  const rows = parseRows(tableInner);
  const body = rows.slice(headerRows(rows, hasHeader));
  const keys = body.map((r) => sortKey(r.cells[nameCol] ?? "")).filter(Boolean);
  for (let i = 1; i < keys.length; i++) if (keys[i] < keys[i - 1]) return false;
  return true;
}
