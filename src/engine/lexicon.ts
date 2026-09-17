// Словарь языка нормативных правок.
//
// Смысл файла — вынести ЗНАНИЕ О ФОРМУЛИРОВКАХ из кода в данные. Одну и ту же
// правку пишут по-разному: «изложить в следующей редакции», «изложить в новой
// редакции», «читать в следующей редакции», «сформулировать следующим образом».
// Если каждую формулировку описывать отдельным правилом в коде, поддержка
// превращается в бесконечную гонку. Здесь же достаточно добавить основу слова
// в нужный список — разбор подхватит её сам.
//
// Слова сравниваются по ОСНОВЕ: падежные окончания русского языка отбрасывать
// перечислением бессмысленно («фразой», «фразу», «фразами», «фразе»…), поэтому
// основа задаётся до окончания, а совпадение проверяется как «начинается с».

/** Что делает инструкция. */
export type Action = "replace" | "add" | "delete" | "substitute";

/** Над чем именно она это делает. */
export type ObjectKind =
  | "point"
  | "heading"
  | "sentence"
  | "word"
  | "paragraph"
  | "footnote"
  | "row"
  | "preamble"
  | "appendix";

/** Где именно относительно якоря. */
export type Position = "after" | "before" | "end" | "start";

export interface StemEntry<T> {
  /** Основа слова без окончания, в нижнем регистре. */
  stem: string;
  value: T;
}

/**
 * Глаголы действия.
 *
 * «Заменить» намеренно отнесён к substitute: в нормативных текстах он почти
 * всегда означает «слово X заменить на Y». Замена пункта целиком пишется через
 * «изложить», и если рядом окажется объект-пункт, разбор это учтёт.
 */
export const ACTION_VERBS: StemEntry<Action>[] = [
  { stem: "излож", value: "replace" },
  { stem: "изложи", value: "replace" },
  { stem: "сформулиров", value: "replace" },
  { stem: "читать", value: "replace" },
  { stem: "дополн", value: "add" },
  { stem: "добав", value: "add" },
  { stem: "включ", value: "add" },
  { stem: "исключ", value: "delete" },
  { stem: "удал", value: "delete" },
  { stem: "отмен", value: "delete" },
  { stem: "замен", value: "substitute" },
];

/** «Признать утратившим силу» — тот же смысл, что «исключить». */
export const DELETE_PHRASES = [
  "признать утратившим силу",
  "признать утратившими силу",
  "считать утратившим силу",
];

/** Объекты правки. */
export const OBJECTS: StemEntry<ObjectKind>[] = [
  { stem: "пункт", value: "point" },
  { stem: "подпункт", value: "point" },
  { stem: "предложени", value: "sentence" },
  { stem: "абзац", value: "paragraph" },
  { stem: "слов", value: "word" },
  { stem: "фраз", value: "word" },
  { stem: "цифр", value: "word" },
  { stem: "числ", value: "word" },
  { stem: "текст", value: "word" },
  { stem: "наименовани", value: "heading" },
  { stem: "заголов", value: "heading" },
  { stem: "формулировк", value: "word" },
  { stem: "предлог", value: "word" },
  { stem: "сноск", value: "footnote" },
  { stem: "строк", value: "row" },
  { stem: "преамбул", value: "preamble" },
  { stem: "приложени", value: "appendix" },
];

/** Окончания творительного падежа: ими помечают ТО, ЧЕМ дополняют. */
export const INSTRUMENTAL_ENDINGS = ["ом", "ем", "ой", "ею", "ью", "ами", "ями", "ом.", "ем."];

/** Указатели места. */
export const POSITIONS: StemEntry<Position>[] = [
  { stem: "после", value: "after" },
  { stem: "перед", value: "before" },
  { stem: "вместо", value: "after" },
];

/** «в конце пункта», «в начале пункта» — место без якорной фразы. */
export const EDGE_POSITIONS: StemEntry<Position>[] = [
  { stem: "в конце", value: "end" },
  { stem: "в начале", value: "start" },
];

/** Порядковые числительные для предложений и абзацев (−1 — последнее). */
export const ORDINALS: StemEntry<number>[] = [
  { stem: "перв", value: 1 },
  { stem: "втор", value: 2 },
  { stem: "трет", value: 3 },
  { stem: "четверт", value: 4 },
  { stem: "пят", value: 5 },
  { stem: "шест", value: 6 },
  { stem: "последн", value: -1 },
];

/**
 * Формулировки «в такой-то редакции». Нужны не для смысла (действие уже задано
 * глаголом), а чтобы отличить директиву от обычного предложения текста.
 */
export const REDACTION_MARKERS = [
  "следующей редакции",
  "новой редакции",
  "редакции",
  "следующего содержания",
  "следующим образом",
];

/**
 * Правка «по тексту»: относится ко всему документу, а не к пункту. Такие
 * инструкции опасно применять как точечные — их надо отличать явно.
 */
export const GLOBAL_MARKERS = ["по тексту", "по всему тексту", "во всем тексте", "во всём тексте"];

/** Пометки о перенумерации. */
export const RENUMBER_POINTS = [
  "перенумерацией пунктов",
  "перенумерацией последующих",
  "изменением нумерации",
  "нумерации последующих",
];
export const RENUMBER_FOOTNOTES = ["перенумерацией сносок", "и сносок", "нумерации сносок"];

/** Слово начинается с основы (регистр не важен). */
export function hasStem(word: string, stem: string): boolean {
  return word.toLowerCase().startsWith(stem);
}

/** Найти в тексте все вхождения основ словаря с позициями. */
export function findStems<T>(
  text: string,
  dict: StemEntry<T>[],
): { value: T; index: number; word: string; stem: string }[] {
  const out: { value: T; index: number; word: string; stem: string }[] = [];
  const wordRe = /[А-Яа-яЁёA-Za-z]+\.?/g;
  let m: RegExpExecArray | null;
  while ((m = wordRe.exec(text)) !== null) {
    const word = m[0];
    // Из нескольких подходящих основ берём самую длинную: «подпункт» точнее,
    // чем «пункт», а «формулировк» — чем «слов».
    let best: StemEntry<T> | null = null;
    for (const entry of dict) {
      if (hasStem(word, entry.stem) && (!best || entry.stem.length > best.stem.length)) {
        best = entry;
      }
    }
    if (best) out.push({ value: best.value, index: m.index, word, stem: best.stem });
  }
  return out;
}

/** Стоит ли слово в творительном падеже («пунктом», «предложением», «фразой»). */
export function isInstrumental(word: string): boolean {
  const w = word.toLowerCase().replace(/\.$/, "");
  return INSTRUMENTAL_ENDINGS.some((e) => w.endsWith(e.replace(/\.$/, "")));
}

/** Есть ли в тексте любая из фраз. */
export function containsAny(text: string, phrases: string[]): boolean {
  const low = text.toLowerCase();
  return phrases.some((p) => low.includes(p));
}
