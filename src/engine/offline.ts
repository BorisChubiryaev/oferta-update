// Детерминированный парсер инструкций из документов «Изменения».
//
// Формулировки в нормативных документах устойчивы, но не единообразны: один и
// тот же смысл записывают то «Пункт 4.8 изложить в следующей редакции», то
// «изложить п. 4.8 в следующей редакции», то «Изложить п. 7.7.2. Оферты…».
// Поэтому разбор идёт не одним всеобъемлющим шаблоном, а набором правил,
// которые пробуются по очереди; всё, что не распозналось, помечается как
// требующее ручной обработки — молча терять правки нельзя, оператор должен
// увидеть каждую строку исходного документа.
import { ACTION_VERBS, OBJECTS } from "./lexicon";
import { parseInstruction, type Ctx, type Draft } from "./parse-instruction";
import { normalizeQuotes, type DocBlock } from "./text";
import type { Operation } from "./types";

let idc = 0;
function nid(src: string) {
  idc += 1;
  return `${src}#${idc}`;
}
export function resetIds() {
  idc = 0;
}

/**
 * Предел склейки абзацев в одну инструкцию.
 *
 * Новая редакция преамбулы — это весь список Ключевых Компаний информационного
 * партнерства, по абзацу на компанию, и правки как раз добавляют туда новые
 * компании: любой «с запасом» подобранный предел рано или поздно обрежет текст
 * на середине списка, а обрезанная редакция молча заменит преамбулу огрызком.
 * Поэтому предел держим высоким — склейку и без него останавливают незакрытая
 * кавычка и начало следующей директивы.
 */
const MAX_MERGE = 60;

function tidy(s: string): string {
  return s.replace(/\u00A0/g, " ").trim().replace(/\s+/g, " ");
}

/** Опечатки, встречающиеся в исходных документах, чинятся до разбора. */
function fixTypos(s: string): string {
  return s
    .replace(/(\d)\.\.(\d)/g, "$1.$2") // «п.6..10»
    .replace(/следующе\s+редакции/gi, "следующей редакции");
}

/** Абзац вместе с таблицами, которые идут сразу за ним. */
interface Para {
  text: string;
  tablesAfter: string[][][];
}

/** Инструкция (возможно, из нескольких абзацев) и её «своя» таблица. */
interface Unit {
  text: string;
  tablesAfter: string[][][];
  tablesBefore: string[][][];
}

/** Баланс «ёлочек» в строке: >0 — кавычка осталась незакрытой. */
function quoteDepth(s: string): number {
  let d = 0;
  for (const ch of s) {
    if (ch === "«") d++;
    else if (ch === "»") d--;
  }
  return d;
}

/**
 * Инструкция часто занимает несколько абзацев: директива заканчивается
 * двоеточием, а новая редакция идёт отдельными абзацами. Признак продолжения —
 * незакрытая «ёлочка» (перечень ККИП в преамбуле тянется на пять абзацев) либо
 * двоеточие в конце и «ёлочка» в начале следующего абзаца.
 */
function mergeUnits(paras: Para[]): Unit[] {
  const out: Unit[] = [];
  for (let i = 0; i < paras.length; i++) {
    const first = i;
    let cur = paras[i].text;
    let last = i;
    let taken = 0;
    while (i + 1 < paras.length && taken < MAX_MERGE) {
      const next = paras[i + 1].text;
      const unclosed = quoteDepth(cur) > 0;
      const colonThenQuote = /:$/.test(cur) && /^«/.test(next);
      if (!unclosed && !colonThenQuote) break;
      // Незакрытая кавычка в одном абзаце (в документах это бывает) не должна
      // засасывать весь остаток документа: следующая директива важнее.
      if (unclosed && !colonThenQuote && startsNewDirective(next)) break;
      // Границу абзацев сохраняем переводом строки: новая редакция преамбулы —
      // это список компаний по абзацу на компанию, и склеив их пробелом мы
      // получили бы вместо списка один нечитаемый абзац на две тысячи знаков.
      cur = cur + "\n" + next;
      i++;
      last = i;
      taken++;
    }
    out.push({
      text: cur,
      tablesAfter: paras[last].tablesAfter,
      // Таблицы, стоящие ПЕРЕД инструкцией, — запасной вариант: в исходниках
      // встречается перепутанная вёрстка, когда содержимое правки набрано выше
      // своей директивы (так в Изменениях № 76 с «Дополнить таблицу в п.4»).
      tablesBefore: first > 0 ? paras[first - 1].tablesAfter : [],
    });
  }
  return out;
}

/**
 * Абзац начинается как новая директива, а не как продолжение редакции.
 *
 * Проверка нужна из-за незакрытых кавычек в исходных документах: без неё
 * «Сноску 32 после слов «ООО СК «Сбербанк Страхование» дополнить…» тянет за
 * собой следующую инструкцию, и обе правки достаются одной сноске.
 *
 * Признак — первое слово абзаца: либо глагол-директива, либо объект правки
 * (пункт, сноска, абзац, преамбула). Список берётся из словаря, а не из
 * отдельного перечня, чтобы расширялся вместе с ним.
 */
function startsNewDirective(text: string): boolean {
  // Правки часто идут нумерованным списком («3. Внести изменение в …»). Без
  // снятия номера абзац не опознавался как новая директива: незакрытая кавычка
  // предыдущей правки проглатывала весь остаток документа, и правки 3, 4, …
  // пропадали из результата целиком.
  const body = text.replace(/^\s*\d+[.)]\s+/, "");
  if (/^(?:в\s+раздел|в\s+приложени|внести)/i.test(body)) return true;
  const firstWord = body.match(/^[А-Яа-яЁёA-Za-z]+/);
  if (!firstWord) return false;
  const w = firstWord[0].toLowerCase();
  return (
    ACTION_VERBS.some((v) => w.startsWith(v.stem)) || OBJECTS.some((o) => w.startsWith(o.stem))
  );
}

// ── контекст разбора ────────────────────────────────────────────────────────

/**
 * Заголовок раздела/приложения: задаёт контекст для следующих строк, сам
 * правкой не является. Отличаем по отсутствию глагола-директивы: «В разделе 6
 * «ПРАВА…»:» — заголовок, а «В п.6.7 … дополнить фразой» — уже правка.
 */
function asContextLine(text: string, ctx: Ctx): boolean {
  if (/(изложить|дополнить|исключить|заменить|добавить|удалить)/i.test(text)) return false;
  // Заголовки идут нумерованным списком («2. В подпункте 7.4 …:»), поэтому
  // номер списка снимаем — шаблоны ниже привязаны к началу строки.
  const body = text.replace(/^\s*\d+[.)]\s*/, "");
  // «В подпункте 7.4 раздела 7 «ПЕРСОНАЛЬНЫЕ ДАННЫЕ» (…):» — адрес для идущих
  // следом строк-тире. Без этого номер пункта назывался только здесь и
  // терялся, а каждое тире становилось правкой без адреса — «формулировка не
  // распознана».
  const sub = body.match(
    /^В\s+(?:под)?пункт[а-яё]*\s*№?\s*(\d+(?:\.\d+)*)\.?(?:\s+раздел[а-яё]*\s*№?\s*(\d+))?[^»]*(?:«([^»]*)»)?[^:]*:$/i,
  );
  if (sub) {
    ctx.subPoint = sub[1];
    if (sub[2]) ctx.section = sub[2];
    if (sub[3]) ctx.sectionTitle = sub[3];
    return true;
  }
  const sec = body.match(/^В\s+раздел[а-я]*\s+(\d+)\s*(?:«([^»]*)»)?/i);
  if (sec) {
    ctx.section = sec[1];
    ctx.sectionTitle = sec[2] ?? undefined;
    ctx.appendix = undefined;
    ctx.subPoint = undefined;
    return true;
  }
  const app = body.match(/^В\s+приложени[а-я]*\s*№?\s*(\d+)/i);
  if (app) {
    // «В Приложении 7» — это сама Оферта, а не приложение внутри неё: контекст
    // вложенного приложения тут сбрасывается, иначе все последующие пункты
    // начали бы искаться в несуществующем «приложении 7» Оферты.
    ctx.appendix = app[1] === "7" ? undefined : app[1];
    ctx.section = undefined;
    return true;
  }
  return false;
}

// ── сборка операций ─────────────────────────────────────────────────────────

/**
 * Переводы строки нужны ровно одной цели — преамбуле, где новая редакция
 * состоит из нескольких абзацев. Во всех остальных правках текст ложится в
 * ОДИН абзац, и оставшийся в нём перевод строки был бы мусором внутри
 * предложения; якорь и искомая фраза ищутся по тексту абзаца, поэтому в них
 * переводов строки быть не может тем более.
 */
function flatten(s: string | undefined): string | undefined {
  return s === undefined ? undefined : s.replace(/\s+/g, " ").trim();
}

function toOperation(d: Draft, text: string, sourceDoc: string): Operation {
  const keepParagraphs = d.target.kind === "preamble";
  return {
    id: nid(sourceDoc),
    sourceDoc,
    type: d.type,
    target: d.target,
    anchor: flatten(d.anchor),
    find: flatten(d.find),
    payload: keepParagraphs ? d.payload : flatten(d.payload),
    sentenceIndex: d.sentenceIndex,
    paragraphIndex: d.paragraphIndex,
    rows: d.rows,
    rowNumbers: d.rowNumbers,
    rowRange: d.rowRange,
    note: d.note,
    renumberFootnotes: /перенумерац\w*\s+сносок|и\s+сносок/i.test(text),
    renumberPoints: /перенумерац\w*\s+пункт|изменением\s+нумерации/i.test(text),
    rawText: flatten(text)!,
    confidence: d.confidence,
    warnings: d.warnings,
  };
}


// ── табличные правила ───────────────────────────────────────────────────────
//
// Правки приложений разбираются отдельно: их «текст» лежит не в кавычках, а в
// таблице документа «Изменения», и слотовый разбор тут не помощник.

type TableRule = (text: string, ctx: Ctx, tables: string[][][]) => Draft[] | null;

function appendixIn(text: string): string | null {
  const m = text.match(/приложени[а-я]*\s*№?\s*(\d+)/i);
  return m ? m[1] : null;
}

/**
 * Изложить в новой редакции СТРОКИ таблицы приложения: «Изложить пункты 1 и 2
 * таблицы в п.3 Приложения № 2 … в следующей редакции:» + таблица с новыми
 * строками.
 *
 * Здесь «пункты 1 и 2» — это НОМЕРА СТРОК таблицы, а «п.3» — пункт, в котором
 * та таблица лежит. Без такого разбора правка попадала в общую ветку «одна
 * редакция на несколько пунктов» и уходила оператору целиком, хотя новые
 * строки прямо под инструкцией.
 */
const ruleAppendixRows: TableRule = (text, ctx, tables) => {
  if (!/приложени/i.test(text)) return null;
  if (!/следующ(?:ей\s+редакции|его\s+содержания)/i.test(text)) return null;
  // Номера строк называют перед словом «таблиц»: «пункты 1 и 2 таблицы…»,
  // «пункты 17, 24 таблицы…», «строки 3, 4 таблицы…».
  const listed = text.match(/(?:пункт|строк)[а-яё]*\s+((?:\d+\s*(?:,|и)\s*)*\d+)\s*таблиц/i);
  if (!listed) return null;
  const rowNumbers = (listed[1].match(/\d+/g) ?? []).map((n) => parseInt(n, 10));
  if (!rowNumbers.length) return null;
  const appendix = appendixIn(text) ?? ctx.appendix ?? "?";
  // Пункт, В КОТОРОМ лежит таблица: «таблицы в п.3», «таблицы п. 3».
  const tp = text.match(/таблиц[а-яё]*\s*(?:в\s+)?п\.?\s*(\d+)/i);
  const want = new Set(rowNumbers);
  let rows: string[][] = [];
  for (const tbl of tables) {
    const hit = tbl.filter((r) => want.has(parseInt((r[0] || "").trim(), 10)));
    if (hit.length > rows.length) rows = hit.map((r) => r.map((c) => c.trim()));
  }
  return [
    {
      type: "replace_table_rows",
      target: { kind: "appendix_table", appendix, point: tp ? tp[1] : undefined },
      rows,
      rowNumbers,
      confidence: rows.length ? 0.8 : 0.4,
      warnings: rows.length ? undefined : ["новые данные строк не найдены в документе"],
    },
  ];
};

/** Изложить Приложение N в алфавитном порядке. */
const ruleSortAlpha: TableRule = (text) => {
  if (!/алфавитн/i.test(text) || !/приложени/i.test(text) || /Дополнить/i.test(text)) return null;
  return [
    {
      type: "sort_table_alpha",
      target: { kind: "appendix_table", appendix: appendixIn(text) ?? "1" },
      confidence: 0.8,
    },
  ];
};

/**
 * Дополнить Приложение N пунктом/пунктами следующего содержания (строки
 * таблицы). Число дополняемых пунктов заранее неизвестно — «пунктом»,
 * «пунктами», «пункты»: перечислять формы бессмысленно, важно лишь, что речь о
 * содержании, которое лежит в таблице под инструкцией.
 */
const ruleAppendixNewRow: TableRule = (text, _ctx, tables) => {
  if (!/Дополнить\s+Приложени/i.test(text)) return null;
  if (!/пункт[а-яё]*\s+следующего\s+содержания/i.test(text)) return null;
  const appendix = appendixIn(text) ?? "1";
  const rows = companyRows(tables);
  return [
    {
      type: /алфавитн/i.test(text) ? "insert_table_row_alpha" : "append_table_rows",
      target: { kind: "appendix_table", appendix },
      rows,
      confidence: rows.length ? 0.75 : 0.4,
      warnings: rows.length ? undefined : ["данные новой строки не найдены в документе"],
    },
  ];
};

/**
 * Строки-данные из таблиц документа «Изменения».
 *
 * Берём таблицу, в которой большинство строк похожи на записи реестра —
 * наименование организации либо номер по порядку. Так отсекаются служебные
 * таблицы шапки ВНД («Реквизиты ВНД», «История ВНД»), которые есть в каждом
 * документе и раньше могли подмениться содержимым правки.
 */
function companyRows(tables: string[][][]): string[][] {
  let best: string[][] = [];
  for (const tbl of tables) {
    const rows = tbl
      .map((r) => r.map((c) => c.trim()))
      .filter((r) => r.some((c) => c) && r.length >= 2);
    if (rows.length < 1) continue;
    const dataLike = rows.filter(
      (r) => r.some((c) => /(ООО|ОАО|АО|АНО|ПАО|НПФ)\b|«/.test(c)) || /^\d+$/.test(r[0]),
    );
    if (dataLike.length >= Math.ceil(rows.length / 2) && dataLike.length > best.length) {
      best = rows;
    }
  }
  return best;
}

/**
 * Внести изменения в таблицу Приложения N … добавить следующие строки (+
 * таблица с новыми строками сразу под инструкцией).
 */
const ruleAppendixAddRows: TableRule = (text, ctx, tables) => {
  if (!/приложени/i.test(text) || !/таблиц/i.test(text)) return null;
  // Падеж не перечисляем: «следующие строки», «следующими строками».
  if (!/(?:добавить|дополнить)\s+(?:ниже)?следующ[а-яё]*\s+строк/i.test(text)) return null;
  const rows = companyRows(tables);
  return [
    {
      type: "append_table_rows",
      target: { kind: "appendix_table", appendix: appendixIn(text) ?? ctx.appendix ?? "2" },
      rows,
      confidence: rows.length ? 0.8 : 0.4,
      warnings: rows.length ? undefined : ["новые строки таблицы не найдены в документе"],
    },
  ];
};

/**
 * Правка ЗАГОЛОВКОВ столбцов таблицы приложения.
 *
 * Движок умеет добавлять и заменять строки данных, но не перестраивать шапку
 * таблицы: у неё объединённые ячейки и своё оформление, и подмена её строкой
 * текста испортила бы таблицу. Поэтому правку отдаём оператору — но с точным
 * указанием, что именно и где менять, вместо «формулировка не распознана».
 */
const ruleAppendixColumnTitles: TableRule = (text, ctx, tables) => {
  if (!/наименовани[яй]\s+столбц/i.test(text)) return null;
  const appendix = appendixIn(text) ?? ctx.appendix ?? "?";
  // Новая редакция заголовков лежит в таблице под инструкцией. Раньше движок
  // сообщал «поправьте вручную», но НЕ показывал, что именно вносить, и
  // оператору приходилось открывать исходный документ и искать это там.
  const rows = firstRows(tables);
  return [
    {
      type: "manual",
      target: { kind: "appendix_table", appendix },
      rows,
      note:
        `заголовки столбцов таблицы Приложения № ${appendix} нужно заменить вручную — ` +
        "у шапки объединённые ячейки" +
        (rows.length ? ". Новая редакция заголовков показана ниже" : ""),
      confidence: 0.5,
      warnings: ["заголовки столбцов таблицы не заменяются автоматически"],
    },
  ];
};

/** Первая непустая таблица под инструкцией — её содержимое и есть правка. */
function firstRows(tables: string[][][]): string[][] {
  for (const tbl of tables) {
    const rows = tbl.map((r) => r.map((c) => c.trim())).filter((r) => r.some((c) => c));
    if (rows.length) return rows;
  }
  return [];
}

/**
 * Дополнить ТАБЛИЦУ в п. K Приложения N строкой (строками) следующего
 * содержания.
 *
 * Без этого правила «Дополнить таблицу в п.4 Приложения № 2 … строкой
 * следующего содержания:» разбиралась как добавление обычного ТЕКСТОВОГО
 * пункта с номером 4 — в Оферту молча добавлялся несуществующий пункт.
 */
const ruleAppendixTableNewRows: TableRule = (text, ctx, tables) => {
  if (!/(?:дополнить|добавить)\s+таблиц/i.test(text)) return null;
  if (!/строк[а-яё]*\s+следующего\s+содержания/i.test(text)) return null;
  const appendix = appendixIn(text) ?? ctx.appendix ?? "?";
  const tp = text.match(/таблиц[а-яё]*\s*(?:в\s+)?п\.?\s*(\d+)/i);
  const rows = firstRows(tables);
  return [
    {
      type: "append_table_rows",
      target: { kind: "appendix_table", appendix, point: tp ? tp[1] : undefined },
      rows,
      confidence: rows.length ? 0.8 : 0.4,
      warnings: rows.length
        ? undefined
        : ["новые строки таблицы не найдены под инструкцией — внесите их вручную"],
    },
  ];
};

/** Дополнить таблицу в п. K Приложения N строками X–Y. */
const ruleAppendRowsRange: TableRule = (text, _ctx, tables) => {
  const m = text.match(
    /Дополнить таблицу[\s\S]*?Приложени[а-я]*\s*№?\s*(\d+)[\s\S]*?строками\s+(\d+)\s*[-–—]\s*(\d+)/i,
  );
  if (!m) return null;
  const from = parseInt(m[2], 10);
  const to = parseInt(m[3], 10);
  let rows: string[][] = [];
  for (const tbl of tables) {
    const hit = tbl.filter((r) => {
      const n = parseInt((r[0] || "").trim(), 10);
      return n >= from && n <= to;
    });
    if (hit.length) {
      rows = hit.map((r) => r.map((c) => c.trim()));
      break;
    }
  }
  return [
    {
      type: "append_table_rows",
      target: { kind: "appendix_table", appendix: m[1], point: "3" },
      rows,
      rowRange: { from, to },
      confidence: rows.length ? 0.8 : 0.4,
      warnings: rows.length ? undefined : ["строки таблицы не найдены в документе"],
    },
  ];
};

const TABLE_RULES: TableRule[] = [
  ruleAppendixColumnTitles,
  ruleAppendixRows,
  ruleAppendixTableNewRows,
  ruleSortAlpha,
  ruleAppendixNewRow,
  ruleAppendixAddRows,
  ruleAppendRowsRange,
];

/**
 * Правка адресована не Оферте, а другому разделу Альбома форм.
 *
 * «П.1.1 и п. 1.2 раздела «Общие положения» Альбома форм дополнить…» — это
 * правка самого Альбома, а в Оферте пункт 1.1 тоже есть, и разбор молча
 * создавал там пункт с чужим текстом. Признак — упоминание Альбома форм БЕЗ
 * упоминания Оферты или приложения к ней.
 */
function aimedOutsideOffer(text: string): boolean {
  return /альбом[а-яё]*\s+форм/i.test(text) && !/оферт|приложени/i.test(text);
}

/** Похоже ли, что абзац вообще содержит правку (а не шапку документа). */
function looksLikeInstruction(text: string): boolean {
  if (text.length < 15) return false;
  return /(изложить|изложи|дополнить|дополни|исключить|заменить|замени|добавить|добави|удалить|удали|включить|сформулировать|признать утратившим)/i.test(
    text,
  );
}

/**
 * Где начинается собственно перечень правок.
 *
 * Заголовок пишут в обоих порядках: «Внести в Приложение № 7 … следующие
 * изменения:» и «В Приложение 7 … внести следующие изменения:». Привязка к
 * началу абзаца отсекала второй вариант целиком, поэтому ищем сочетание слов
 * где угодно в абзаце.
 *
 * Обязательна ОБЪЯВЛЯЮЩАЯ формула («следующие изменения», «изменения
 * следующего содержания»). Без неё под шаблон «внести» + «изменени» попадала и
 * сама первая правка — «Внести изменения в п.1.2 Приложения № 2 …», — а она
 * вместе с заголовком отбрасывалась: правка молча исчезала из результата, и
 * оператор даже не видел, что её потеряли.
 */
function instructionsStart(paras: string[]): number {
  return paras.findIndex(
    (p) =>
      /внести/i.test(p) &&
      /изменени/i.test(p) &&
      /следующи[ех]\s+изменени|изменени[яй]\s+следующего/i.test(p),
  );
}

/**
 * К чему адресован документ: к Оферте (Приложение 7) или к другому документу.
 *
 * Если заголовка нет вовсе, считаем документ адресованным Оферте. Обратное
 * умолчание опаснее: одна неузнанная строка-заголовок превращала весь документ
 * в список «внесите вручную», и правки молча не применялись. Здесь же ошибка
 * видна сразу — правки просто не найдут своих мест.
 */
function detectScope(header: string | null): { scope: Ctx["scope"]; note?: string } {
  if (header === null) return { scope: "offer" };
  if (/приложени[ея]\s*№?\s*7|оферт/i.test(header)) return { scope: "offer" };
  return {
    scope: "other",
    note:
      "документ изменяет не Оферту (Приложение 7), а другой раздел Альбома форм — " +
      "правку нужно внести вручную в соответствующий документ",
  };
}

/**
 * Результат разбора ОДНОЙ инструкции (единицы) документа «Изменения».
 *
 * Гибридному конвейеру мало плоского списка операций: чтобы доспросить ИИ ровно
 * про то, что детерминированный разбор не осилил, нужна связь «исходная
 * инструкция → что из неё получилось». Плоский список эту связь теряет: две
 * инструкции, давшие по операции `manual`, в нём неразличимы, а склеенная из
 * пяти абзацев редакция преамбулы вообще не восстанавливается из `rawText`
 * операции по одному лишь исходному документу.
 */
export interface UnitResult {
  /** Порядковый номер инструкции в документе (0-based). */
  index: number;
  /** Текст инструкции — уже склеенный, с нормализованными кавычками. */
  text: string;
  /** Таблицы, относящиеся к этой инструкции (после неё, затем — перед ней). */
  tables: string[][][];
  /** Что из инструкции получил детерминированный разбор. */
  operations: Operation[];
}

/** Разбор документа «Изменения» с сохранением связи инструкция → операции. */
export function analyzeChangeDoc(
  blocks: DocBlock[],
  docTables: string[][][],
  sourceDoc: string,
): { units: UnitResult[]; scope: Ctx["scope"] } {
  // Абзацы и «свои» таблицы: таблица приписывается тому абзацу, после которого
  // она идёт в документе, — так правка вида «…добавить следующие строки:» знает,
  // какая именно таблица её, даже если таблиц в документе несколько.
  const raw: Para[] = [];
  for (const b of blocks) {
    if (b.kind === "p") raw.push({ text: b.text, tablesAfter: [] });
    else if (raw.length) raw[raw.length - 1].tablesAfter.push(b.rows);
  }
  // Кавычки приводим к «ёлочкам» по ВСЕМУ документу сразу, а не по абзацам:
  // новая редакция сплошь и рядом занимает несколько абзацев, и открывающая
  // лапка стоит в одном абзаце, а закрывающая — в другом. Поабзацный разбор
  // такую пару найти не может, и текст правки оставался невидимым — правка
  // выглядела как «не найден текст новой редакции». Длину строк normalizeQuotes
  // не меняет, поэтому склейка через \n и обратное разрезание безопасны.
  const texts = normalizeQuotes(raw.map((p) => tidy(p.text)).join("\n"))
    .split("\n")
    .map((p) => fixTypos(p));
  const paras: Para[] = raw
    .map((p, i) => ({ text: texts[i], tablesAfter: p.tablesAfter }))
    .filter((p) => p.text);
  const start = instructionsStart(paras.map((p) => p.text));
  const { scope, note } = detectScope(start >= 0 ? paras[start].text : null);
  const ctx: Ctx = { scope, scopeNote: note };
  const units = mergeUnits(paras.slice(start + 1));

  const results: UnitResult[] = [];
  let unitIndex = 0;
  for (const { text, tablesAfter, tablesBefore } of units) {
    if (asContextLine(text, ctx)) continue;
    if (!looksLikeInstruction(text)) continue;
    const ops: Operation[] = [];
    const pushUnit = (tables: string[][][]) => {
      results.push({ index: unitIndex++, text, tables, operations: ops });
    };

    // Документ (или отдельная правка в нём) адресован не Оферте — не применяем,
    // но и не теряем: оператор увидит строку и внесёт её в нужный документ.
    if (ctx.scope === "other" || aimedOutsideOffer(text)) {
      const note =
        ctx.scope === "other"
          ? ctx.scopeNote
          : "правка относится к другому разделу Альбома форм, а не к Оферте — " +
            "внесите её вручную в соответствующий документ";
      ops.push(
        toOperation(
          {
            type: "manual",
            target: { kind: "point", point: "—" },
            note,
            confidence: 0.6,
            warnings: [note ?? ""],
          },
          text,
          sourceDoc,
        ),
      );
      pushUnit([]);
      continue;
    }

    // Сначала таблицы приложений: их содержимое лежит вне текста инструкции.
    // Своя таблица (та, что идёт сразу под инструкцией) имеет приоритет. Если
    // под инструкцией таблицы нет, берём последнюю, стоящую перед ней (бывает
    // перепутанная вёрстка), и лишь затем — любую таблицу документа.
    const scoped = tablesAfter.length
      ? tablesAfter
      : tablesBefore.length
        ? [tablesBefore[tablesBefore.length - 1]]
        : docTables;
    let drafts: Draft[] | null = null;
    for (const rule of TABLE_RULES) {
      const res = rule(text, ctx, scoped);
      if (res && res.length) {
        drafts = res;
        break;
      }
    }
    // Затем общий разбор по слотам.
    if (!drafts) drafts = parseInstruction(text, ctx);

    if (drafts && drafts.length) {
      for (const d of drafts) ops.push(toOperation(d, text, sourceDoc));
      pushUnit(scoped);
      continue;
    }
    // Ничего не подошло — правка всё равно должна дойти до оператора.
    ops.push(
      toOperation(
        {
          type: "manual",
          target: { kind: "point", point: "—" },
          note: "формулировка не распознана автоматически",
          confidence: 0.3,
          warnings: ["формулировка не распознана — внесите правку вручную"],
        },
        text,
        sourceDoc,
      ),
    );
    pushUnit(scoped);
  }
  return { units: results, scope: ctx.scope };
}

/**
 * Плоский список операций — прежний контракт конвейера.
 *
 * Оставлен как тонкая обёртка над `analyzeChangeDoc`, чтобы у алгоритмического
 * режима (без сети и ИИ) поведение осталось ровно таким, каким было.
 */
export function parseInstructionsOffline(
  blocks: DocBlock[],
  docTables: string[][][],
  sourceDoc: string,
): Operation[] {
  return analyzeChangeDoc(blocks, docTables, sourceDoc).units.flatMap((u) => u.operations);
}
