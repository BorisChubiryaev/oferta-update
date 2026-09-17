// Локация абзаца-цели для операции «изложить в новой редакции».
//
// ПРИНЦИП: находим цель ТОЛЬКО по надёжному признаку и НИКОГДА не «угадываем»
// нечётким совпадением — иначе можно молча испортить не тот пункт. Если
// уверенной цели нет, возвращаем null (движок честно сообщит «не найдено»,
// а оператор увидит это в предпросмотре).
//
// Признаки по приоритету:
//   • термин — абзац, ТЕКСТ которого начинается с имени термина;
//   • пункт/пункт приложения — номер (восстановленный из автонумерации или
//     литеральный) В ПРЕДЕЛАХ нужного раздела/приложения.
import { indexNumberedParagraphs, findByNumber, normNumber } from "./numbering";
import type { NumberedPara } from "./numbering";
import type { Operation } from "./types";

export interface ParaSpan {
  start: number;
  end: number;
  inner: string;
}

function normText(s: string): string {
  return s.replace(/\s+/g, "").toLowerCase();
}
function span(p: NumberedPara): ParaSpan {
  return { start: p.start, end: p.end, inner: p.inner };
}

/** Начало раздела/приложения (для сужения поиска по номеру). */
function regionStart(index: NumberedPara[], op: Operation): number {
  if (op.target.kind !== "appendix_point") return 0;
  // Заголовок приложения — самый надёжный ориентир: нумерация внутри каждого
  // приложения начинается заново, и «п. 1.2» Приложения № 2 обязан искаться
  // строго ниже его заголовка.
  const head = appendixHeading(index, op.target.appendix);
  if (head !== null) return head;
  // Запасной вариант — название приложения, названное в самой инструкции.
  // Он ненадёжен: в «…Приложения № 2 к Приложению 7 «Публичная Оферта…»»
  // в кавычках стоит имя САМОЙ Оферты, поиск приводил к её началу, и пункт
  // 1.2 приложения подменялся пунктом 1.2 основного текста.
  const m = op.rawText.match(/Приложени[а-я]*\s*№?\s*\d+[^«]*«([^»]{6,}?)»/i);
  if (m && !/оферт/i.test(m[1])) {
    const want = normText(m[1]).slice(0, 40);
    const hit = index.find((p) => normText(p.text).includes(want));
    if (hit) return hit.start;
  }
  return 0;
}

/** Позиция заголовка «Приложение № N» в теле Оферты. */
function appendixHeading(index: NumberedPara[], appendix: string): number | null {
  const re = new RegExp(`^приложение\\s*№?\\s*${appendix}(?!\\d)`, "i");
  const hit = index.find((p) => re.test(p.text.trim()));
  return hit ? hit.start : null;
}

/** Литеральный номер в начале абзаца («2.2.», «1.4.») в пределах региона. */
function findByLiteralNumber(index: NumberedPara[], point: string, from: number): ParaSpan | null {
  const want = normNumber(point);
  const re = new RegExp("^\\s*" + want.replace(/\./g, "\\.") + "\\.?(?!\\d)");
  for (const p of index) {
    if (p.start < from) continue;
    if (re.test(p.text)) return span(p);
  }
  return null;
}

/**
 * Смещение заголовка приложения. Нумерация внутри приложений начинается заново,
 * поэтому «п. 2.2» Приложения № 2 и «п. 2.2» основного текста — разные абзацы,
 * и искать пункт приложения надо только ниже его заголовка.
 */
export function appendixOffset(documentXml: string, appendix: string): number {
  for (const anchor of [`Приложение № ${appendix}`, `Приложение №${appendix}`, `Приложение ${appendix}`]) {
    const flat = documentXml.indexOf(anchor);
    if (flat >= 0) return flat;
  }
  return 0;
}

/**
 * Найти абзац пункта по номеру. `from` ограничивает поиск (для приложений —
 * их заголовком), иначе «2.2» приложения совпадёт с термином основного текста.
 */
export function locatePointSpan(
  documentXml: string,
  numberingXml: string | null,
  point: string,
  stylesXml: string | null = null,
  from = 0,
): ParaSpan | null {
  const index = indexNumberedParagraphs(documentXml, numberingXml, stylesXml);
  const byNum = findByNumber(index, point, from);
  if (byNum) return span(byNum);
  return findByLiteralNumber(index, point, from);
}

/**
 * Все абзацы, из которых состоит пункт.
 *
 * Пункт в Оферте — это не один абзац: у «6.1» есть абзац-зачин, оканчивающийся
 * двоеточием, перечисление 1)…2) и завершающий абзац без номера. Инструкция
 * «последнее предложение в п. 6.1» имеет в виду именно завершающий абзац, и
 * без учёта блока правка стирала зачин целиком.
 *
 * Границу определяем по списку и уровню: следующий абзац ТОГО ЖЕ списка с тем
 * же или более высоким уровнем — уже соседний пункт, всё остальное (подпункты
 * другого списка, абзацы без номера) принадлежит текущему.
 */
export function locatePointBlock(
  documentXml: string,
  numberingXml: string | null,
  point: string,
  stylesXml: string | null = null,
  from = 0,
): ParaSpan[] {
  const index = indexNumberedParagraphs(documentXml, numberingXml, stylesXml);
  const want = normNumber(point);
  const i = index.findIndex(
    (p) => p.start >= from && p.number != null && normNumber(p.number) === want,
  );
  if (i >= 0) {
    const head = index[i];
    const out = [head];
    for (let j = i + 1; j < index.length; j++) {
      const p = index[j];
      if (
        p.numId !== null &&
        head.numId !== null &&
        p.numId === head.numId &&
        p.ilvl !== null &&
        head.ilvl !== null &&
        p.ilvl <= head.ilvl
      ) {
        break;
      }
      out.push(p);
    }
    return out.map(span);
  }

  // Резерв: номер не восстановлен автонумерацией — как и locatePointSpan,
  // ищем абзац по началу его ТЕКСТА (документ набран литеральными номерами
  // без w:numPr; так бывает в старых или подготовленных вручную Офертах).
  // Без этого резерва «второй абзац п. 5.1 изложить…» и «дополнить п. 5.1…»
  // в таких документах никогда не находили бы свой пункт — хотя точечная
  // замена всего пункта (locatePointSpan) его прекрасно находит.
  const litRe = new RegExp("^\\s*" + want.replace(/\./g, "\\.") + "\\.?(?!\\d)");
  const litI = index.findIndex((p) => p.start >= from && litRe.test(p.text));
  if (litI < 0) return [];
  const headDepth = want.split(".").length;
  const out = [index[litI]];
  for (let j = litI + 1; j < index.length; j++) {
    const p = index[j];
    // Другой автонумерованный абзац — заведомо иная (не литеральная) часть
    // документа, дальше в неё не заходим.
    if (p.number != null) break;
    const lm = p.text.match(/^\s*(\d+(?:\.\d+)*)\.?(?!\d)/);
    // Абзац с номером ТОЙ ЖЕ или МЕНЬШЕЙ глубины — начало следующего пункта
    // (сосед или раздел выше); с БОЛЬШЕЙ глубиной — подпункт, входит в блок.
    if (lm && lm[1].split(".").length <= headDepth) break;
    out.push(p);
  }
  return out.map(span);
}

/**
 * Последний рубеж: найти пункт по началу его текста.
 *
 * Номера в документах «Изменения» относятся к своей редакции и расходятся с
 * текущей (а в Оферте есть подсписки с зашитым в шаблон номером «7.6.%3»,
 * которые вообще не сдвигаются). Если новая редакция начинается теми же
 * словами, что и действующая, это надёжный признак — но только когда
 * совпадение ЕДИНСТВЕННОЕ: иначе можно переписать не тот пункт.
 */
export function locateByTextPrefix(
  documentXml: string,
  numberingXml: string | null,
  payload: string,
  stylesXml: string | null = null,
  point?: string,
  prefixLength = 60,
): ParaSpan | null {
  const body = payload.replace(/^\s*\d+(?:\.\d+)*\.?\s*/, "");
  const needle = normText(body).slice(0, prefixLength);
  if (needle.length < prefixLength) return null;
  const index = indexNumberedParagraphs(documentXml, numberingXml, stylesXml);
  const hits = index.filter((p) => normText(p.text).startsWith(needle));
  if (hits.length === 1) return span(hits[0]);
  // Совпадений несколько (в Оферте есть однотипные пункты «Для предоставления
  // Сервиса…» в каждом разделе). Тогда доверяем номеру частично: берём
  // ближайшего существующего предка номера и проверяем его текст.
  if (hits.length > 1 && point) {
    const parts = normNumber(point).split(".");
    for (let cut = parts.length - 1; cut >= 1; cut--) {
      const ancestor = findByNumber(index, parts.slice(0, cut).join("."), 0);
      if (ancestor && normText(ancestor.text).startsWith(needle)) return span(ancestor);
    }
  }
  return null;
}

/**
 * Преамбула — вводная часть до раздела 1. Опознаём по устойчивой формуле
 * «публикует/публикуют настоящее предложение заключить договор»: номера у
 * неё нет, а первые слова меняются от редакции к редакции.
 *
 * Преамбула — НЕ один абзац: между именем Банка и словами «публикуют настоящее
 * предложение» перечислены Ключевые Компании информационного партнерства, по
 * абзацу на компанию. Взяв только замыкающий абзац, замена оставила бы старый
 * список компаний на месте и дописала к нему новый — в Оферте оказались бы два
 * списка сразу. Поэтому идём от замыкающего абзаца вверх и захватываем всю
 * вводную часть, останавливаясь на заголовке: нумерованном абзаце или
 * прописной строке-титуле («ПУБЛИЧНАЯ ОФЕРТА»).
 */
export function locatePreamble(
  documentXml: string,
  numberingXml: string | null,
  stylesXml: string | null = null,
): ParaSpan | null {
  const index = indexNumberedParagraphs(documentXml, numberingXml, stylesXml);
  const lastAt = index.findIndex((p) =>
    /публику[ею]т\s+настоящее\s+предложение\s+заключить\s+договор/i.test(p.text),
  );
  if (lastAt < 0) return null;
  let firstAt = lastAt;
  while (firstAt > 0 && !isPreambleBoundary(index[firstAt - 1])) firstAt--;
  return {
    start: index[firstAt].start,
    end: index[lastAt].end,
    inner: index[firstAt].inner,
  };
}

/** Заголовок, выше которого преамбула уже не продолжается. */
function isPreambleBoundary(p: NumberedPara): boolean {
  if (p.number) return true;
  const t = p.text.trim();
  if (!t) return false;
  // Титульные строки набраны прописными («ПРИЛОЖЕНИЕ 7», «ПУБЛИЧНАЯ ОФЕРТА»), а
  // сама преамбула — обычным предложением, поэтому регистр здесь надёжнее
  // любого перечня возможных заголовков.
  const letters = t.replace(/[^А-Яа-яЁёA-Za-z]/g, "");
  return letters.length > 0 && letters === letters.toUpperCase();
}

/**
 * Найти место для ВСТАВКИ нового пункта X (напр. 2.9). Возвращает абзац-якорь
 * и режим: "before" — вставить перед ним (новый станет X), "after" — вставить
 * после предыдущего пункта (когда X ещё нет, добавляем в конец списка).
 * Поиск ведётся в пределах раздела (по номеру раздела перед первой точкой).
 */
export function locatePointInsertion(
  documentXml: string,
  numberingXml: string | null,
  point: string,
  stylesXml: string | null = null,
): { span: ParaSpan; mode: "before" | "after" } | null {
  const index = indexNumberedParagraphs(documentXml, numberingXml, stylesXml);
  const section = normNumber(point).split(".")[0];
  // Регион раздела: заголовок с номером == section (ilvl 0).
  const secHead = section ? findByNumber(index, section, 0) : null;
  const from = secHead ? secHead.start : 0;

  const exact = findByNumber(index, point, from);
  if (exact) return { span: span(exact), mode: "before" };

  // X ещё нет — ищем предыдущий номер того же уровня (X с уменьшённым хвостом).
  const parts = normNumber(point).split(".").map((n) => parseInt(n, 10));
  const last = parts[parts.length - 1];
  for (let p = last - 1; p >= 1; p--) {
    const prevNum = [...parts.slice(0, -1), p].join(".");
    const prev = findByNumber(index, prevNum, from);
    if (prev) return { span: span(prev), mode: "after" };
  }
  return null;
}

/** Найти абзац-цель для замены. Возвращает span либо null (без «угадывания»). */
export function locateReplaceParagraph(
  documentXml: string,
  numberingXml: string | null,
  op: Operation,
  stylesXml: string | null = null,
): ParaSpan | null {
  const index = indexNumberedParagraphs(documentXml, numberingXml, stylesXml);

  // Преамбула — по устойчивой формуле, номера у неё нет.
  if (op.target.kind === "preamble") {
    return locatePreamble(documentXml, numberingXml, stylesXml);
  }

  // Термин ищем СНАЧАЛА по имени: имя — настоящая идентичность термина, а
  // номер в документе «Изменения» относится к своей редакции и легко
  // расходится с текущей. Номер остаётся запасным вариантом.
  if (op.target.kind === "term") {
    if (op.target.term) {
      const want = normText(op.target.term);
      const hit = index.find((p) => normText(p.text).startsWith(want));
      if (hit) return span(hit);
    }
    if (op.target.point) {
      const byNum = findByNumber(index, op.target.point, 0);
      if (byNum) return span(byNum);
      const byLit = findByLiteralNumber(index, op.target.point, 0);
      if (byLit) return byLit;
    }
    return null;
  }

  // Пункт / пункт приложения — по номеру в пределах раздела.
  if (op.target.kind === "point" || op.target.kind === "appendix_point") {
    const point = op.target.point;
    const from = regionStart(index, op);
    const byNum = findByNumber(index, point, from);
    if (byNum) return span(byNum);
    const byLit = findByLiteralNumber(index, point, from);
    if (byLit) return byLit;
  }

  return null;
}
