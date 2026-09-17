// Разбор одной инструкции по слотам.
//
// Вместо набора шаблонов «под каждую формулировку» инструкция раскладывается на
// составляющие: действие (что делаем), объект (над чем), цель (где), якорь
// (относительно каких слов) и текст. Дальше пара «действие × объект» однозначно
// даёт операцию движка.
//
// Такой разбор устойчив к порядку слов («Изложить п. 4.8 в новой редакции» и
// «Пункт 4.8 изложить в следующей редакции» дают одно и то же) и к падежам, а
// расширяется словарём в lexicon.ts, а не новым ветвлением в коде.
//
// Отдельно стоит правило творительного падежа: в русском языке им помечают то,
// ЧЕМ дополняют. «Дополнить пунктом 2.15» — добавляем пункт; «дополнить п. 8.5
// предложением» — добавляем предложение В пункт 8.5. Без этого различия обе
// формулировки выглядят одинаково, а смысл у них противоположный.
import {
  ACTION_VERBS,
  DELETE_PHRASES,
  EDGE_POSITIONS,
  GLOBAL_MARKERS,
  OBJECTS,
  ORDINALS,
  POSITIONS,
  REDACTION_MARKERS,
  RENUMBER_FOOTNOTES,
  RENUMBER_POINTS,
  containsAny,
  findStems,
  isInstrumental,
  type Action,
  type ObjectKind,
  type Position,
} from "./lexicon";
import { extractGuillemet } from "./text";
import type { OpTarget, OpType } from "./types";

export interface Ctx {
  section?: string;
  sectionTitle?: string;
  appendix?: string;
  /**
   * Пункт из строки-заголовка для подпунктов-тире, идущих следом:
   * «В подпункте 7.4 раздела 7 «ПЕРСОНАЛЬНЫЕ ДАННЫЕ»: — после текста «А»
   * дополнить …; — после текста «Б» дополнить …». Номер пункта назван один раз
   * в заголовке, и без него каждое тире — правка без адреса.
   */
  subPoint?: string;
  scope: "offer" | "other";
  scopeNote?: string;
}

/** Строка-подпункт перечня («— после текста «…» дополнить …»). */
export function isSubItem(text: string): boolean {
  return /^\s*[-–—•]\s*/.test(text);
}

export interface Draft {
  type: OpType;
  target: OpTarget;
  anchor?: string;
  find?: string;
  payload?: string;
  sentenceIndex?: number;
  paragraphIndex?: number;
  rows?: string[][];
  rowNumbers?: number[];
  rowRange?: { from: number; to: number };
  confidence: number;
  warnings?: string[];
  note?: string;
}

interface Quote {
  text: string;
  start: number;
  end: number;
}

/**
 * \u041F\u0435\u0440\u0435\u0432\u043E\u0434\u044B \u0441\u0442\u0440\u043E\u043A\u0438 \u0432\u043D\u0443\u0442\u0440\u0438 \u043A\u0430\u0432\u044B\u0447\u0435\u043A \u0421\u041E\u0425\u0420\u0410\u041D\u042F\u042E\u0422\u0421\u042F: \u043E\u043D\u0438 \u043E\u0442\u043C\u0435\u0447\u0430\u044E\u0442 \u0433\u0440\u0430\u043D\u0438\u0446\u044B \u0430\u0431\u0437\u0430\u0446\u0435\u0432
 * \u043C\u043D\u043E\u0433\u043E\u0430\u0431\u0437\u0430\u0446\u043D\u043E\u0439 \u0440\u0435\u0434\u0430\u043A\u0446\u0438\u0438 (\u0441\u043F\u0438\u0441\u043E\u043A \u043A\u043E\u043C\u043F\u0430\u043D\u0438\u0439 \u0432 \u043F\u0440\u0435\u0430\u043C\u0431\u0443\u043B\u0435). \u0421\u0445\u043B\u043E\u043F\u044B\u0432\u0430\u0435\u0442\u0441\u044F \u0442\u043E\u043B\u044C\u043A\u043E
 * \u0433\u043E\u0440\u0438\u0437\u043E\u043D\u0442\u0430\u043B\u044C\u043D\u044B\u0439 \u043F\u0440\u043E\u0431\u0435\u043B.
 */
function tidy(s: string): string {
  return s
    .replace(/\u00A0/g, " ")
    .split("\n")
    .map((line) => line.replace(/[^\S\n]+/g, " ").trim())
    .join("\n")
    .trim();
}

/**
 * Признак того, что «сбалансированная» кавычка захватила лишнее: внутри неё
 * оказалась следующая директива. Так бывает из-за опечаток в исходных
 * документах — «Сноску 32 после слов «ООО СК «Сбербанк Страхование» дополнить»:
 * внешняя кавычка не закрыта, и сбалансированный разбор проглатывает весь
 * остаток инструкции вместе с новой редакцией.
 */
const SWALLOWED_DIRECTIVE =
  /(дополн(?:ить|ив)|добав(?:ить|ив)|излож(?:ить|ив)|замен(?:ить|ив)|исключ(?:ить|ив)|удал(?:ить|ив))\s+[^«»]{0,40}«/i;

/**
 * Кавычки «…» с позициями. Сначала пробуем сбалансированный разбор (он нужен
 * для вложенных кавычек вроде «(в Системе «Сбербанк Онлайн»)»), а если он
 * захватил директиву — откатываемся к первой закрывающей кавычке.
 */
function allQuotes(text: string): Quote[] {
  const out: Quote[] = [];
  let from = 0;
  for (;;) {
    const open = text.indexOf("«", from);
    if (open < 0) break;
    const g = extractGuillemet(text, open);
    if (!g) break;
    let end = g.endIndex;
    let content = g.content;
    // Кавычка не закрыта вовсе (в исходниках это обычное дело — «…Приложения 7
    // "Публичная Оферта … "Удобный доступ" Альбома форм…»): сбалансированный
    // разбор дотягивает её до конца строки и проглатывает вместе с ней новую
    // редакцию. Ограничиваемся первой закрывающей кавычкой.
    if (!g.balanced || SWALLOWED_DIRECTIVE.test(content)) {
      const firstClose = text.indexOf("»", open + 1);
      if (firstClose > open) {
        end = firstClose;
        content = text.slice(open + 1, firstClose);
      }
    }
    out.push({ text: tidy(content), start: open, end });
    from = end + 1;
  }
  return out;
}

interface PointRef {
  num: string;
  index: number;
}

/**
 * Номера пунктов с позициями. Перечисления вида «п. 1, 2.2 и 2.3» дают
 * несколько номеров: после маркера «п.»/«пункт» подхватываются и последующие
 * номера, пока они идут через запятую или «и».
 */
function findPoints(text: string): PointRef[] {
  const out: PointRef[] = [];
  // Сокращение пишут по-разному: «п. 4.10», «п.4.10», «п 4.10», «пп. 5.1».
  // Проверка «слева не буква» отсекает случайные «п» внутри слов.
  const re = /(?:подпункт[а-яё]*|пункт[а-яё]*|(?<![А-Яа-яЁё])пп?\.?\s*)\s*(\d+(?:\.\d+)*)/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    out.push({ num: m[1].replace(/\.$/, ""), index: m.index });
    // Хвост перечисления: «, 2.2 и 2.3.»
    let tail = re.lastIndex;
    const chain = /^\s*(?:,|и)\s*(\d+(?:\.\d+)*)/;
    for (;;) {
      const rest = text.slice(tail);
      const c = rest.match(chain);
      if (!c) break;
      out.push({ num: c[1].replace(/\.$/, ""), index: tail });
      tail += c[0].length;
    }
    re.lastIndex = Math.max(re.lastIndex, tail);
  }
  return out;
}

/**
 * Номер Оферты в Альбоме форм: Оферта — это САМО Приложение 7, а не приложение
 * внутри неё. Поэтому «п. 9.6 Приложения 7 «Публичная Оферта…»» адресует
 * обычный пункт Оферты, а не вложенное приложение; искать в Оферте приложение
 * с номером 7 бессмысленно — пункт не находился вовсе.
 */
const OFFER_APPENDIX = "7";

/**
 * Номер приложения ВНУТРИ Оферты. Ссылки на саму Оферту («Приложения 7»)
 * пропускаются, поэтому «п.1.2 Приложения № 2 к Приложению 7» даёт 2 — то
 * вложенное приложение, о котором и говорит правка.
 */
function appendixIn(text: string): string | null {
  const re = /приложени[а-я]*\s*№?\s*(\d+)/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    if (m[1] !== OFFER_APPENDIX) return m[1];
  }
  return null;
}

/**
 * Диапазон номеров пунктов через дефис («5.1-5.3», «7.6–7.7»), который
 * findPoints НЕ раскрывает (она понимает лишь перечисление через запятую или
 * «и»). Без этой проверки «Пункты 7.6-7.7 изложить в редакции: …» тихо
 * применялась бы только к пункту 7.6, а 7.7 терялся без единого предупреждения.
 */
function hasUnexpandedRange(text: string, points: PointRef[]): boolean {
  return points.some((p) => {
    // p.index — начало ВСЕГО совпадения findPoints (со слова «пункт»), а не
    // самой цифры; ищем цифру от этой позиции, чтобы проверить символ сразу
    // после номера.
    const at = text.indexOf(p.num, p.index);
    if (at < 0) return false;
    const after = text.slice(at + p.num.length);
    return /^\s*[-–—]\s*\d+(?:\.\d+)*/.test(after);
  });
}

/** Номер раздела: «раздела 5», «разделе 7». */
function sectionNumberIn(text: string): string | null {
  const m = text.match(/раздел[а-яё]*\s*№?\s*(\d+)/i);
  return m ? m[1] : null;
}

function footnoteNumberIn(text: string): number | null {
  const m = text.match(/сноск[а-я]*\s*№?\s*(\d+)/i);
  return m ? parseInt(m[1], 10) : null;
}

/** Имя термина из новой редакции («Стороны – совместное упоминание…»). */
function termName(payload: string): string | null {
  const body = payload.replace(/^\s*\d+(?:\.\d+)*\.?\s*/, "").trim();
  const dash = body.search(/\s[–—-]\s/);
  if (dash <= 0 || dash > 140) return null;
  return body.slice(0, dash).trim();
}

function targetForPoint(point: string, ctx: Ctx, text: string, payload?: string): OpTarget {
  const app = ctx.appendix ?? (/приложени/i.test(text) ? appendixIn(text) : null);
  if (app) return { kind: "appendix_point", appendix: app, point };
  const inTerms =
    point.startsWith("2.") || ctx.section === "2" || /термин/i.test(ctx.sectionTitle ?? "");
  if (inTerms && payload) {
    const term = termName(payload);
    if (term) return { kind: "term", section: ctx.section, point, term };
  }
  return {
    kind: "point",
    section: ctx.section ?? point.split(".")[0],
    point,
    heading: ctx.sectionTitle,
  };
}

// ── слоты ───────────────────────────────────────────────────────────────────

interface Slots {
  action: Action | null;
  actionAt: number;
  /** Объект сразу после глагола — он и задаёт, над чем работаем. */
  object: ObjectKind | null;
  objectAt: number;
  objectWord: string;
  objectInstrumental: boolean;
  ordinal: number | null;
  position: Position | null;
  positionAt: number;
  points: PointRef[];
  quotes: Quote[];
  /** Все найденные объекты — для правил, которым важен не только главный. */
  allObjects: { value: ObjectKind; index: number; word: string }[];
}

/**
 * Текст директивы без содержимого кавычек.
 *
 * Внутри «…» лежит НОВАЯ РЕДАКЦИЯ — это содержание, а не грамматика правки.
 * Если искать слова-подсказки по всему тексту, редакция «…персонализации
 * предложений.» превратит замену пункта в замену предложения. Позиции
 * символов сохраняем, заменяя содержимое пробелами.
 */
function maskQuotes(text: string, quotes: Quote[]): string {
  const chars = Array.from(text);
  for (const q of quotes) {
    for (let i = q.start; i <= q.end && i < chars.length; i++) chars[i] = " ";
  }
  return chars.join("");
}

function readSlots(text: string): Slots {
  const quotes = allQuotes(text);
  const grammar = maskQuotes(text, quotes);
  const verbs = findStems(grammar, ACTION_VERBS);
  const objects = findStems(grammar, OBJECTS);
  const ordinals = findStems(grammar, ORDINALS);
  const positions = findStems(grammar, POSITIONS);

  let action: Action | null = verbs.length ? verbs[0].value : null;
  let actionAt = verbs.length ? verbs[0].index : -1;
  if (!action && containsAny(text, DELETE_PHRASES)) {
    action = "delete";
    actionAt = 0;
  }
  // «Дополнить пунктом 5.6 и изложить в следующей редакции» — ведёт первый
  // глагол: он определяет, что вообще делаем, остальные лишь уточняют.

  // Какой из объектов главный. Порядок предпочтений выведен из того, как
  // формулируют правки:
  //  1) объект, за которым сразу идёт кавычка («слова «X» исключить») — правка
  //     явно о нём;
  //  2) первый объект ПОСЛЕ глагола («дополнить предложением…»);
  //  3) первый объект вообще — когда глагол в конце («Сноску к пункту 5.3
  //     изложить»): подлежащее стоит в начале.
  const quotedObject = objects.find((o) =>
    quotes.some((q) => q.start > o.index && q.start - o.index <= o.word.length + 2),
  );
  const after = objects.filter((o) => o.index > actionAt);
  // При дополнении главный объект — тот, что в творительном падеже: «Дополнить
  // пункт 3.4 абзацем» добавляет абзац, а не пункт. Без этого предпочтения
  // побеждает «пункт», который лишь указывает адрес правки.
  const instrumental =
    action === "add" ? objects.find((o) => o.index > actionAt && isInstrumental(o.word)) : undefined;
  let chosen = instrumental ?? quotedObject ?? after[0] ?? objects[0] ?? null;

  // Порядковое числительное относится к объекту, если стоит прямо перед ним.
  // Такой объект главнее прочих: во фразе «Последнее предложение пункта 6.1
  // изложить…» речь о предложении, а «пункт» лишь указывает, где его искать.
  let ordinal: number | null = null;
  let ordinalObject: (typeof objects)[number] | null = null;
  for (const o of ordinals) {
    // «второй абзац» и «абзац второй» — обе формы обычны, поэтому смотрим по
    // обе стороны и выбираем ближайшее подходящее существительное.
    const near = objects
      .filter(
        (x) =>
          Math.abs(x.index - o.index) <= 20 &&
          (x.value === "sentence" || x.value === "paragraph"),
      )
      .sort((a, b) => Math.abs(a.index - o.index) - Math.abs(b.index - o.index))[0];
    if (near) {
      ordinal = o.value;
      ordinalObject = near;
      break;
    }
  }

  // Творительный падеж сильнее порядкового числительного: во фразе «в конце
  // второго предложения дополнить словами» дополняем словами, а «второе
  // предложение» лишь указывает место.
  if (ordinalObject && !instrumental) chosen = ordinalObject;

  // Указатель места учитываем только если за ним идёт кавычка-якорь.
  let position: Position | null = null;
  let positionAt = -1;
  for (const p of positions) {
    const q = quotes.find((x) => x.start > p.index && x.start - p.index <= 40);
    if (q) {
      position = p.value;
      positionAt = p.index;
      break;
    }
  }
  if (!position) {
    for (const e of EDGE_POSITIONS) {
      const at = grammar.toLowerCase().indexOf(e.stem);
      if (at >= 0) {
        position = e.value;
        positionAt = at;
        break;
      }
    }
  }

  return {
    action,
    actionAt,
    object: chosen ? chosen.value : null,
    objectAt: chosen ? chosen.index : -1,
    objectWord: chosen ? chosen.word : "",
    objectInstrumental: chosen ? isInstrumental(chosen.word) : false,
    ordinal,
    position,
    positionAt,
    points: findPoints(grammar),
    quotes,
    allObjects: objects.map((o) => ({ value: o.value, index: o.index, word: o.word })),
  };
}

// ── сборка операций ─────────────────────────────────────────────────────────

/**
 * Пары «после слов «A» дополнить «B»». Их в одной инструкции может быть
 * несколько: «после слов «X» дополнить словами «Y», после слова «Z» дополнить
 * предлогом «с»» — раньше в текст пункта уезжала вся фраза целиком.
 */
function insertionPairs(
  text: string,
  quotes: Quote[],
): { position: Position; anchor: string; payload: string }[] {
  const pairs: { position: Position; anchor: string; payload: string }[] = [];
  // Между указателем места и кавычкой встречается разное — «после слов»,
  // «после фразы», «после цифр», «после слов пункта». Перечислять формы
  // бессмысленно: достаточно короткого промежутка без кавычек.
  const re = /(после|перед)\s+[^«»]{0,40}(?=«)/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    const anchor = quotes.find((q) => q.start >= m!.index);
    if (!anchor) continue;
    const rest = text.slice(anchor.end + 1);
    // Так же и после глагола: «дополнить словами», «дополнить фразой»,
    // «дополнить следующими компаниями:» — важно лишь, что дальше идёт кавычка.
    const verb = rest.match(/^\s*,?\s*(?:дополнить|добавить|включить)\s+[^«»]{0,90}(?=«)/i);
    if (!verb) continue;
    const payload = quotes.find((q) => q.start > anchor.end);
    if (!payload) continue;
    pairs.push({
      position: m[1].toLowerCase() === "перед" ? "before" : "after",
      anchor: anchor.text,
      payload: payload.text,
    });
    re.lastIndex = payload.end;
  }
  return pairs;
}

/**
 * Кавычка с НОВОЙ РЕДАКЦИЕЙ — первая после оборота «изложить в следующей
 * редакции:» / «следующего содержания:».
 *
 * Брать просто последнюю кавычку нельзя: в заголовке правки нередко стоит
 * название документа («…Приложения 7 «Публичная Оферта…» Альбома форм…»), и при
 * непарных кавычках в исходнике именно оно оказывалось последним — в пункт
 * вместо новой редакции уезжало название Оферты.
 */
function redactionPayload(text: string, quotes: Quote[]): Quote | null {
  const m = text.match(
    /(?:следующей|новой|указанной)\s+редакции|следующего\s+содержания|следующим\s+образом/i,
  );
  if (!m || m.index === undefined) return null;
  return quotes.find((q) => q.start > m.index!) ?? null;
}

/**
 * Разложить общую новую редакцию по пунктам, которых она касается.
 *
 * «Пункты 1.1, 1.2, 1.3 … изложить в следующей редакции: «1.1 … 1.2 … 1.3 …»» —
 * одна кавычка на три пункта. Сам текст редакции при этом пронумерован, и
 * границы кусков заданы в нём явно: ищем каждый номер в начале строки или
 * предложения и режем по найденным позициям.
 *
 * Возвращает null, если хотя бы один номер не найден или они идут не по
 * порядку: тогда границы кусков неизвестны, и разносить правку наугад нельзя —
 * лучше честно отдать её оператору целиком.
 */
function splitByPointNumbers(payload: string, points: string[]): string[] | null {
  if (points.length < 2 || !payload) return null;
  const starts: number[] = [];
  let from = 0;
  for (const p of points) {
    const re = new RegExp(`(?:^|[\\n.;)]\\s*|\\s)(${p.replace(/\./g, "\\.")})\\.?(?=[\\s)]|$)`, "g");
    re.lastIndex = from;
    const m = re.exec(payload);
    if (!m) return null;
    const at = m.index + m[0].indexOf(m[1]);
    if (starts.length && at <= starts[starts.length - 1]) return null;
    starts.push(at);
    from = at + m[1].length;
  }
  return starts.map((at, i) =>
    payload.slice(at, i + 1 < starts.length ? starts[i + 1] : payload.length).trim(),
  );
}

/**
 * Номера пунктов, которыми размечен сам текст новой редакции («1.1 … 1.2 …»).
 * Нужны, чтобы раскрыть диапазон «7.6-7.7»: инструкция границы не называет, а
 * текст редакции — называет. Считаем разметкой только номера того же уровня,
 * что и начало диапазона, идущие строго по возрастанию.
 */
function pointsNamedIn(payload: string, first: string): string[] | null {
  const depth = first.split(".").length;
  const re = /(?:^|[\n.;)]\s*|\s)(\d+(?:\.\d+)*)\.?(?=[\s)]|$)/g;
  const out: string[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(payload)) !== null) {
    const num = m[1];
    if (num.split(".").length !== depth) continue;
    if (out.length === 0 ? num !== first : compareNums(num, out[out.length - 1]) <= 0) continue;
    out.push(num);
  }
  return out.length >= 2 ? out : null;
}

function compareNums(a: string, b: string): number {
  const x = a.split(".").map(Number);
  const y = b.split(".").map(Number);
  for (let i = 0; i < Math.max(x.length, y.length); i++) {
    const d = (x[i] ?? 0) - (y[i] ?? 0);
    if (d) return d;
  }
  return 0;
}

const OP_BY_ADD_OBJECT: Partial<Record<ObjectKind, OpType>> = {
  point: "insert_point",
  sentence: "append_sentence",
  paragraph: "append_paragraph",
  footnote: "add_footnote",
  // «Пункт 6.2 дополнить словами «…»» без указания места — текст идёт в конец
  // пункта; случай «после слов …» перехватывается раньше.
  word: "append_sentence",
};

/**
 * Разобрать инструкцию. Возвращает набор операций (в одной строке их может быть
 * несколько) либо null, если строка на правку не похожа.
 */
export function parseInstruction(text: string, ctx: Ctx): Draft[] | null {
  const s = readSlots(text);
  if (!s.action) return null;

  const renumberPoints = containsAny(text, RENUMBER_POINTS);
  const renumberFootnotes = containsAny(text, RENUMBER_FOOTNOTES);
  void renumberPoints;
  void renumberFootnotes;

  const lastQuote = s.quotes.length ? s.quotes[s.quotes.length - 1] : null;
  const pointNums = s.points.map((p) => p.num);
  // Свой номер важнее унаследованного; номер из заголовка подхватывают только
  // строки-тире, чтобы контекст не «протёк» в следующую самостоятельную правку.
  const firstPoint = pointNums[0] ?? (isSubItem(text) ? (ctx.subPoint ?? null) : null);

  // ── сноски ────────────────────────────────────────────────────────────────
  // «дополнить сноской» — про сноску, даже если главным объектом выбрано слово
  // из оборота «после слов».
  // «дополнить сноской» (творительный) и «добавив сноску к п. 9.3.1 «…»»
  // (винительный) — одно и то же действие. Без второй формы правка уезжала в
  // ветку «дополнить пункт» и вместо сноски создавала ФИКТИВНЫЙ пункт 9.3.1 с
  // текстом сноски внутри.
  const footnoteAdd = s.allObjects.some(
    (o) => o.value === "footnote" && (isInstrumental(o.word) || o.index > s.actionAt),
  );
  if (s.object === "footnote" || footnoteAdd) {
    const num = footnoteNumberIn(text);
    if (s.action === "add" && (s.objectInstrumental || footnoteAdd)) {
      const pair = insertionPairs(text, s.quotes)[0];
      const contentIdx = text.search(/содержания|редакции/i);
      // Якорь пишут не только через «после слов»: встречается «сноску к словам
      // «X» дополнить сноской следующего содержания: «Y»». Если пары не нашлось,
      // но кавычек две — первая это слова-якорь, вторая текст сноски.
      const payload =
        (contentIdx >= 0 ? s.quotes.find((q) => q.start > contentIdx)?.text : undefined) ??
        lastQuote?.text ??
        "";
      const anchor =
        pair?.anchor ??
        (s.quotes.length >= 2 && s.quotes[0].text !== payload ? s.quotes[0].text : "");
      return [
        {
          type: "add_footnote",
          target: firstPoint
            ? targetForPoint(firstPoint, ctx, text)
            : { kind: "point", point: "?" },
          anchor,
          payload,
          confidence: anchor && payload ? 0.8 : 0.4,
          warnings: anchor ? undefined : ["не найдены слова-якорь для сноски"],
        },
      ];
    }
    if (s.action === "replace" && (num !== null || firstPoint)) {
      return [
        {
          type: "replace_footnote",
          target:
            num !== null
              ? { kind: "footnote", number: num }
              : { kind: "footnote", number: 0, atPoint: firstPoint! },
          payload: lastQuote?.text ?? "",
          confidence: lastQuote ? (num !== null ? 0.85 : 0.7) : 0.4,
          warnings:
            num === null ? ["сноска задана пунктом — будет взята первая сноска этого пункта"] : undefined,
        },
      ];
    }
    if (s.action === "delete" && (num !== null || firstPoint)) {
      return [
        {
          type: "delete_footnote",
          target:
            num !== null
              ? { kind: "footnote", number: num }
              : { kind: "footnote", number: 0, atPoint: firstPoint! },
          confidence: 0.8,
        },
      ];
    }
    if (num !== null && s.action === "add") {
      const pair = insertionPairs(text, s.quotes)[0];
      if (pair) {
        return [
          {
            type: "insert_after",
            target: { kind: "footnote", number: num },
            anchor: pair.anchor,
            payload: ", " + pair.payload.replace(/\.$/, "") + ".",
            confidence: 0.85,
          },
        ];
      }
    }
  }

  // ── преамбула ─────────────────────────────────────────────────────────────
  // Преамбулу правят и глаголом «изложить», и оборотом «дополнить … и указать в
  // следующей редакции»: во втором случае действие разбирается как «дополнить»,
  // но в кавычках всё равно лежит ПОЛНЫЙ новый текст преамбулы — списком новых
  // компаний его не дополняют, его переписывают целиком. Отличаем по обороту «в
  // такой-то редакции»: без него «дополнить преамбулу словами» остаётся
  // обычным дополнением.
  if (s.object === "preamble" && (s.action === "replace" || s.action === "add")) {
    const quote = redactionPayload(text, s.quotes) ?? lastQuote;
    if (s.action === "replace" || (quote && containsAny(text, REDACTION_MARKERS))) {
      return [
        {
          type: "replace",
          target: { kind: "preamble" },
          payload: quote?.text ?? "",
          confidence: quote ? 0.85 : 0.4,
          warnings: quote ? undefined : ["не найден текст новой редакции преамбулы"],
        },
      ];
    }
  }

  // ── вставка относительно якоря (может быть несколько в одной инструкции) ──
  if (s.action === "add" && (s.position === "after" || s.position === "before")) {
    const pairs = insertionPairs(text, s.quotes);
    // Вставлять можно и в пункт, и в сноску — «Сноску 5 после слов … дополнить».
    const fnNum = footnoteNumberIn(text);
    const target: OpTarget | null = firstPoint
      ? targetForPoint(firstPoint, ctx, text)
      : fnNum !== null
        ? { kind: "footnote", number: fnNum }
        : null;
    if (pairs.length && target) {
      const intoFootnote = target.kind === "footnote";
      return pairs.map((p) => ({
        type: (p.position === "before" ? "insert_before" : "insert_after") as OpType,
        target,
        anchor: p.anchor,
        // Пробел ставим со стороны, обращённой к якорю: при вставке ПЕРЕД
        // словами он нужен справа, иначе получится «В любой моментБанк вправе».
        payload: intoFootnote
          ? ", " + p.payload.replace(/\.$/, "") + "."
          : p.position === "before"
            ? p.payload + " "
            : " " + p.payload,
        confidence: 0.85,
      }));
    }
  }

  // ── замена по всему тексту ────────────────────────────────────────────────
  // «По тексту Оферты слова «А» заменить словами «Б»» — правка не привязана к
  // пункту, и применять её как точечную нельзя.
  if (s.action === "substitute" && containsAny(text, GLOBAL_MARKERS) && s.quotes.length >= 2) {
    const find = s.quotes[0];
    const to = s.quotes.find((q) => q.start > find.end);
    if (to) {
      return [
        {
          type: "replace_words_global",
          target: { kind: "preamble" },
          find: find.text,
          payload: to.text,
          confidence: 0.8,
          warnings: ["правка применяется по всему тексту — проверьте число замен"],
        },
      ];
    }
  }

  // ── удаление слов ─────────────────────────────────────────────────────────
  if (s.action === "delete" && s.object === "word") {
    // Кавычка со словами стоит либо после глагола («удалить слова «X»»), либо
    // до него («слова «X» исключить») — берём ближайшую к слову-объекту.
    const findQuote =
      s.quotes.find((q) => q.start > s.actionAt) ?? s.quotes.find((q) => q.start > s.objectAt);
    const anchorQuote = s.position ? s.quotes.find((q) => q.start > s.positionAt) : undefined;
    if (findQuote && firstPoint) {
      return [
        {
          type: "delete_words",
          target: targetForPoint(firstPoint, ctx, text),
          find: findQuote.text,
          anchor: anchorQuote && anchorQuote !== findQuote ? anchorQuote.text : undefined,
          confidence: 0.85,
        },
      ];
    }
  }

  // ── замена слов ───────────────────────────────────────────────────────────
  if (s.action === "substitute" && s.quotes.length >= 2 && pointNums.length) {
    const find = s.quotes[0];
    const to = s.quotes.find((q) => q.start > find.end);
    if (to) {
      return pointNums.map((point) => ({
        type: "replace_words" as OpType,
        target: targetForPoint(point, ctx, text),
        find: find.text,
        payload: to.text,
        confidence: 0.85,
      }));
    }
  }

  // ── исключение абзаца пункта ──────────────────────────────────────────────
  if (s.action === "delete" && s.object === "paragraph" && firstPoint) {
    return [
      {
        type: "delete_paragraph",
        target: targetForPoint(firstPoint, ctx, text),
        paragraphIndex: s.ordinal ?? -1,
        confidence: 0.8,
      },
    ];
  }

  // ── исключение сноски ─────────────────────────────────────────────────────
  if (s.action === "delete" && s.object === "footnote") {
    const num = footnoteNumberIn(text);
    if (num !== null) {
      return [
        {
          type: "delete_footnote",
          target: { kind: "footnote", number: num },
          confidence: 0.8,
        },
      ];
    }
  }

  // ── исключение пункта ─────────────────────────────────────────────────────
  if (s.action === "delete" && (s.object === "point" || (!s.object && firstPoint))) {
    // Диапазон «пункты 5.1–5.3» раскрыть нельзя, не зная, какие номера есть в
    // документе, а угадывать здесь опаснее, чем отдать оператору.
    const range = text.match(/(\d+(?:\.\d+)*)\s*[-–—]\s*(\d+(?:\.\d+)*)/);
    if (range) {
      return [
        {
          type: "manual",
          target: targetForPoint(range[1], ctx, text),
          note: `исключение диапазона пунктов ${range[1]}–${range[2]}: перечислите пункты по отдельности или удалите их вручную`,
          confidence: 0.4,
          warnings: ["диапазон пунктов не раскрывается автоматически"],
        },
      ];
    }
    if (firstPoint) {
      return pointNums.map((point) => ({
        type: "delete_point" as OpType,
        target: targetForPoint(point, ctx, text),
        confidence: 0.85,
      }));
    }
  }

  // ── дополнение ────────────────────────────────────────────────────────────
  if (s.action === "add") {
    const payload = (redactionPayload(text, s.quotes) ?? lastQuote)?.text ?? "";
    // Творительный падеж указывает, ЧЕМ дополняем.
    if (s.objectInstrumental && s.object) {
      const type = OP_BY_ADD_OBJECT[s.object];
      if (type === "insert_point") {
        // «Дополнить пунктом 2.15»: номер после объекта — номер НОВОГО пункта.
        const newPoint =
          s.points.find((p) => p.index > s.actionAt)?.num ?? firstPoint ?? "?";
        return [
          {
            type,
            target: targetForPoint(newPoint, ctx, text, payload),
            payload,
            confidence: payload ? 0.8 : 0.4,
            warnings: payload ? undefined : ["не найден текст нового пункта"],
          },
        ];
      }
      if (type && firstPoint) {
        return [
          {
            type,
            target: targetForPoint(firstPoint, ctx, text),
            payload,
            // «в конце второго предложения» — дополняем именно его, а не пункт.
            sentenceIndex: type === "append_sentence" ? (s.ordinal ?? undefined) : undefined,
            confidence: payload ? 0.85 : 0.4,
          },
        ];
      }
    }
    // Без творительного падежа «дополнить п. 7.3 … в следующей редакции»
    // означает добавление самого пункта 7.3.
    if (firstPoint) {
      return [
        {
          type: "insert_point",
          target: targetForPoint(firstPoint, ctx, text, payload),
          payload,
          confidence: payload ? 0.8 : 0.4,
          warnings: payload ? undefined : ["не найден текст нового пункта"],
        },
      ];
    }
  }

  // ── изложение в новой редакции ────────────────────────────────────────────
  if (s.action === "replace" || s.action === "substitute") {
    const payload = (redactionPayload(text, s.quotes) ?? lastQuote)?.text ?? "";
    // Заголовок раздела — обычный абзац документа, пронумерованный на верхнем
    // уровне: «раздел 5» находится по номеру «5».
    if (s.object === "heading") {
      const section = sectionNumberIn(text) ?? firstPoint;
      if (!section) return null;
      return [
        {
          type: "replace",
          target: { kind: "point", section, point: section },
          payload,
          confidence: payload ? 0.75 : 0.4,
          warnings: ["правка заголовка — проверьте, что найден именно он"],
        },
      ];
    }
    if (!firstPoint) return null;
    // Несколько пунктов и одна редакция. Если текст редакции сам пронумерован,
    // границы кусков в нём заданы явно — разносим по пунктам. Если нет, какую
    // часть текста к какому пункту относить, из инструкции не следует.
    //
    // Проверять здесь s.object === "point" нельзя: во фразе «п.п. 1.1, 1.2, 1.3
    // Приложения № 2 …» главным объектом оказывается «приложение» (сокращение
    // «п.п.» словарь объектов не видит), правка проваливалась в общую ветку
    // ниже и применялась ТОЛЬКО к первому номеру — остальные пункты молча
    // терялись. Исключаем лишь предложение и абзац: там номер пункта — адрес,
    // а не перечень целей.
    if (pointNums.length > 1 && s.object !== "sentence" && s.object !== "paragraph") {
      const parts = splitByPointNumbers(payload, pointNums);
      if (parts) {
        return pointNums.map((point, i) => ({
          type: "replace" as OpType,
          target: targetForPoint(point, ctx, text, parts[i]),
          payload: parts[i],
          confidence: 0.7,
          warnings: [
            `общая редакция для пунктов ${pointNums.join(", ")} разнесена по номерам из её текста — сверьте границы`,
          ],
        }));
      }
      return [
        {
          type: "manual",
          target: targetForPoint(firstPoint, ctx, text),
          note: `новая редакция сразу для пунктов ${pointNums.join(", ")}: разнесите правки по пунктам вручную`,
          confidence: 0.4,
          warnings: ["одна редакция на несколько пунктов не разбирается автоматически"],
        },
      ];
    }
    // Диапазон через дефис («Пункты 7.6-7.7 изложить…») findPoints не
    // раскрывает: без этой проверки правка тихо применилась бы только к
    // первому номеру диапазона, а остальные терялись бы бесследно.
    if (s.object === "point" && hasUnexpandedRange(text, s.points)) {
      // Какие именно номера входят в диапазон, по инструкции неизвестно — зато
      // их называет сам текст редакции, если он пронумерован.
      const spanned = pointsNamedIn(payload, firstPoint);
      const parts = spanned && splitByPointNumbers(payload, spanned);
      if (spanned && parts) {
        return spanned.map((point, i) => ({
          type: "replace" as OpType,
          target: targetForPoint(point, ctx, text, parts[i]),
          payload: parts[i],
          confidence: 0.7,
          warnings: [
            `диапазон раскрыт по номерам из текста редакции (${spanned.join(", ")}) — сверьте границы`,
          ],
        }));
      }
      return [
        {
          type: "manual",
          target: targetForPoint(firstPoint, ctx, text),
          note: `новая редакция для диапазона пунктов начиная с ${firstPoint}: диапазон не раскрывается автоматически, разнесите правки по пунктам вручную`,
          confidence: 0.4,
          warnings: ["диапазон пунктов в правке не раскрывается автоматически"],
        },
      ];
    }
    if (s.object === "sentence") {
      return [
        {
          type: "replace_sentence",
          target: targetForPoint(firstPoint, ctx, text),
          sentenceIndex: s.ordinal ?? -1,
          payload,
          confidence: payload ? 0.85 : 0.4,
        },
      ];
    }
    if (s.object === "paragraph") {
      return [
        {
          type: "replace_paragraph",
          target: targetForPoint(firstPoint, ctx, text),
          paragraphIndex: s.ordinal ?? -1,
          payload,
          confidence: payload ? 0.85 : 0.4,
        },
      ];
    }
    return pointNums.slice(0, 1).map((point) => ({
      type: "replace" as OpType,
      target: targetForPoint(point, ctx, text, payload),
      payload,
      confidence: payload ? 0.85 : 0.4,
      warnings: payload ? undefined : ["не найден текст новой редакции"],
    }));
  }

  return null;
}
