// Низкоуровневые операции над OOXML (WordprocessingML).
//
// Правки Оферты выполняются ХИРУРГИЧЕСКИ над сырой XML-строкой document.xml /
// footnotes.xml, без полного разбора и пересборки дерева. Так исходный
// юридический документ (сотни КБ) сохраняется байт-в-байт всюду, кроме
// точечно изменённых мест.

/** Экранирование текста для вставки в <w:t>. */
export function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * Обратное преобразование XML-сущностей в текст.
 *
 * Числовые ссылки (&#34; и &#x22;) обязательны: Word записывает прямую кавычку
 * именно так, и без их разбора весь текст в прямых кавычках («Публичная Оферта»
 * в кавычках-лапках) для парсера инструкций просто не существует — правка
 * выглядит как «не найден текст новой редакции».
 *
 * &amp; раскрывается последним, иначе «&amp;#34;» превратился бы в кавычку.
 */
export function decodeXml(s: string): string {
  return s
    .replace(/&#(\d+);/g, (_, d: string) => safeCodePoint(parseInt(d, 10)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h: string) => safeCodePoint(parseInt(h, 16)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function safeCodePoint(code: number): string {
  if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return "";
  try {
    return String.fromCodePoint(code);
  } catch {
    return "";
  }
}

/** rPr для «выделения цветом» — красный шрифт EE0000, как в образце Оферты. */
export function colorRunProps(color = "EE0000"): string {
  return `<w:rPr><w:color w:val="${color}"/></w:rPr>`;
}

/** Собрать один текстовый run с заданными свойствами. */
export function makeRun(text: string, rPr = ""): string {
  return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
}

/** Токен-run с его границами и текстом. */
interface RunToken {
  start: number; // индекс начала <w:r> в xml
  end: number; // индекс сразу после </w:r>
  full: string;
  text: string; // декодированный видимый текст рана (конкатенация w:t)
  simple: boolean; // ровно один <w:t> и нет иных текстовых узлов (можно точно резать)
  rPr: string; // сырой <w:rPr>…</w:rPr> или ""
  tAttrs: string; // атрибуты тега <w:t …>
}

const RUN_RE = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;
const WT_RE = /<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
const RPR_RE = /^<w:r(?:\s[^>]*)?>(\s*<w:rPr>[\s\S]*?<\/w:rPr>)?/;

function tokenizeRuns(xml: string): RunToken[] {
  const tokens: RunToken[] = [];
  let m: RegExpExecArray | null;
  RUN_RE.lastIndex = 0;
  while ((m = RUN_RE.exec(xml)) !== null) {
    const full = m[0];
    const start = m.index;
    const end = start + full.length;
    const texts: string[] = [];
    let wt: RegExpExecArray | null;
    WT_RE.lastIndex = 0;
    let tAttrs = "";
    let wtCount = 0;
    while ((wt = WT_RE.exec(full)) !== null) {
      texts.push(decodeXml(wt[2]));
      tAttrs = wt[1] || "";
      wtCount++;
    }
    // «Сложные» узлы, мешающие точной резке (табы, разрывы, картинки).
    const hasOther = /<w:(tab|br|drawing|object|footnoteReference|endnoteReference)\b/.test(full);
    const rprMatch = full.match(RPR_RE);
    const rPr = rprMatch && rprMatch[1] ? rprMatch[1] : "";
    tokens.push({
      start,
      end,
      full,
      text: texts.join(""),
      simple: wtCount === 1 && !hasOther,
      rPr,
      tAttrs,
    });
  }
  return tokens;
}

/** Позиция символа: индекс рана и смещение внутри его текста. */
interface CharPos {
  runIdx: number;
  offInRun: number;
}

/**
 * Найти якорную фразу в тексте набора ранов, игнорируя ВСЕ пробелы
 * (в документах встречается «СК «Сбербанк» и «СК«Сбербанк» — без пробела).
 * Возвращает позицию ПОСЛЕДНЕГО символа якоря либо null.
 */
function isIgnorable(ch: string): boolean {
  // Пробелы и кавычки-«ёлочки» игнорируем при сопоставлении: в инструкции
  // « » служат разделителями фразы, а в тексте — ещё и кавычками, и эти
  // роли часто совпадают на одном символе.
  return /\s/.test(ch) || ch === "«" || ch === "»";
}

/** Диапазон фразы в ранах: от первого символа до последнего включительно. */
interface PhraseRange {
  from: CharPos;
  to: CharPos;
}

/** Найти фразу в тексте ранов, игнорируя пробелы и «ёлочки». */
function findPhrase(tokens: RunToken[], phrase: string, occurrence = 0): PhraseRange | null {
  const chars: { ch: string; pos: CharPos }[] = [];
  tokens.forEach((t, runIdx) => {
    for (let i = 0; i < t.text.length; i++) {
      const ch = t.text[i];
      if (isIgnorable(ch)) continue;
      chars.push({ ch: ch.toLowerCase(), pos: { runIdx, offInRun: i } });
    }
  });
  const flat = chars.map((c) => c.ch).join("");
  const needle = Array.from(phrase)
    .filter((ch) => !isIgnorable(ch))
    .join("")
    .toLowerCase();
  if (!needle) return null;
  let idx = -1;
  for (let n = 0; n <= occurrence; n++) {
    idx = flat.indexOf(needle, idx + 1);
    if (idx < 0) return null;
  }
  return { from: chars[idx].pos, to: chars[idx + needle.length - 1].pos };
}

/**
 * Номер вхождения `find`, которое идёт ПОСЛЕ фразы `anchor`.
 *
 * Инструкции вида «после слов A удалить слова B» опираются на порядок: сами
 * по себе слова B могут встречаться в пункте несколько раз, и без привязки к A
 * мы вырежем не то место.
 */
export function phraseOccurrenceAfter(
  xml: string,
  anchor: string,
  find: string,
): number | null {
  const flatten = (t: string) =>
    Array.from(t)
      .filter((ch) => !isIgnorable(ch))
      .join("")
      .toLowerCase();
  const tokens = tokenizeRuns(xml);
  const text = flatten(tokens.map((t) => t.text).join(""));
  const a = flatten(anchor);
  const f = flatten(find);
  if (!f) return null;
  const anchorAt = a ? text.indexOf(a) : 0;
  if (anchorAt < 0) return null;
  const target = text.indexOf(f, anchorAt + (a ? a.length : 0));
  if (target < 0) return null;
  let n = 0;
  for (let i = text.indexOf(f); i >= 0 && i < target; i = text.indexOf(f, i + 1)) n++;
  return n;
}

/** Сколько раз фраза встречается в ранах. */
export function countPhrase(xml: string, phrase: string): number {
  const tokens = tokenizeRuns(xml);
  let n = 0;
  while (findPhrase(tokens, phrase, n)) n++;
  return n;
}

function findAnchorEnd(tokens: RunToken[], anchor: string): CharPos | null {
  const r = findPhrase(tokens, anchor);
  return r ? r.to : null;
}

/** Собрать ран с тем же оформлением и заданным куском текста. */
function sliceRun(tok: RunToken, from: number, to: number): string {
  const text = tok.text.slice(from, to);
  if (!text) return "";
  const tAttrs = tok.tAttrs || ' xml:space="preserve"';
  return `<w:r>${tok.rPr}<w:t${tAttrs}>${escapeXml(text)}</w:t></w:r>`;
}

/**
 * Заменить фразу `find` на готовые раны `newRuns`.
 *
 * Фраза может быть разорвана на несколько ранов (Word режет текст произвольно),
 * поэтому режем крайние раны и выбрасываем те, что целиком внутри фразы. Раны
 * со сносками/картинками внутри диапазона — стоп-сигнал: вырезав их, мы молча
 * потеряем ссылку на сноску, поэтому в таком случае честно отказываемся.
 */
export function replacePhraseRuns(
  xml: string,
  find: string,
  newRuns: string,
  occurrence = 0,
): InsertResult {
  const tokens = tokenizeRuns(xml);
  const range = findPhrase(tokens, find, occurrence);
  if (!range) {
    return { xml, ok: false, message: `фраза не найдена: «${find}»`, orderKey: -1 };
  }
  const a = tokens[range.from.runIdx];
  const b = tokens[range.to.runIdx];
  for (let i = range.from.runIdx; i <= range.to.runIdx; i++) {
    if (!tokens[i].simple) {
      return {
        xml,
        ok: false,
        message: `фраза «${find}» пересекает сноску или объект — правка требует ручной обработки`,
        orderKey: -1,
      };
    }
  }
  const head = sliceRun(a, 0, range.from.offInRun);
  const tail = sliceRun(b, range.to.offInRun + 1, b.text.length);
  const rebuilt = head + newRuns + tail;
  return {
    xml: xml.slice(0, a.start) + rebuilt + xml.slice(b.end),
    ok: true,
    message: "фраза заменена",
    orderKey: a.start,
  };
}

export interface InsertResult {
  xml: string;
  ok: boolean;
  message: string;
  /** Позиция вставки в исходном xml (для сортировки по порядку Оферты). */
  orderKey: number;
}

/**
 * Вставить готовые раны `newRuns` СРАЗУ ПОСЛЕ якорной фразы `anchor`.
 * Ран, где заканчивается якорь, при необходимости разрезается, чтобы
 * выделение встало ровно после нужных слов.
 */
export function insertAfterAnchor(
  xml: string,
  anchor: string,
  newRuns: string,
): InsertResult {
  const tokens = tokenizeRuns(xml);
  const end = findAnchorEnd(tokens, anchor);
  if (!end) {
    return { xml, ok: false, message: `якорь не найден: «${anchor}»`, orderKey: -1 };
  }
  const tok = tokens[end.runIdx];
  let cut = end.offInRun + 1; // режем сразу после последнего символа якоря
  // Если дальше (через пробелы) стоит закрывающая », вставка должна встать
  // ПОСЛЕ неё: «…Страхование»‹сюда›. Но сами пробелы забирать нельзя — иначе
  // они уедут в «левую» часть, и получится «заключения  Банком и ККИПДоговора».
  let look = cut;
  while (look < tok.text.length && /\s/.test(tok.text[look])) look++;
  if (tok.text[look] === "»") cut = look + 1;

  if (tok.simple) {
    const before = tok.text.slice(0, cut);
    const after = tok.text.slice(cut);
    const tAttrs = tok.tAttrs || ' xml:space="preserve"';
    const rebuilt =
      `<w:r>${tok.rPr}<w:t${tAttrs}>${escapeXml(before)}</w:t></w:r>` +
      newRuns +
      (after.length
        ? `<w:r>${tok.rPr}<w:t${tAttrs}>${escapeXml(after)}</w:t></w:r>`
        : "");
    const newXml = xml.slice(0, tok.start) + rebuilt + xml.slice(tok.end);
    return { xml: newXml, ok: true, message: "вставлено после якоря", orderKey: tok.start };
  }

  // Сложный ран — вставляем после него целиком (выделение чуть менее точное).
  const newXml = xml.slice(0, tok.end) + newRuns + xml.slice(tok.end);
  return {
    xml: newXml,
    ok: true,
    message: "вставлено после рана (ран сложный, резка пропущена)",
    orderKey: tok.end,
  };
}

/**
 * Вставить раны ПЕРЕД якорной фразой. Нужно для формулировок «перед словами
 * «X» дополнить словами «Y»» — зеркало insertAfterAnchor.
 */
export function insertBeforeAnchor(
  xml: string,
  anchor: string,
  newRuns: string,
): InsertResult {
  const tokens = tokenizeRuns(xml);
  const range = findPhrase(tokens, anchor);
  if (!range) {
    return { xml, ok: false, message: `якорь не найден: «${anchor}»`, orderKey: -1 };
  }
  const tok = tokens[range.from.runIdx];
  if (!tok.simple) {
    const newXml = xml.slice(0, tok.start) + newRuns + xml.slice(tok.start);
    return { xml: newXml, ok: true, message: "вставлено перед раном", orderKey: tok.start };
  }
  const tAttrs = tok.tAttrs || ' xml:space="preserve"';
  const cut = range.from.offInRun;
  const head = cut ? `<w:r>${tok.rPr}<w:t${tAttrs}>${escapeXml(tok.text.slice(0, cut))}</w:t></w:r>` : "";
  const tail = `<w:r>${tok.rPr}<w:t${tAttrs}>${escapeXml(tok.text.slice(cut))}</w:t></w:r>`;
  const newXml = xml.slice(0, tok.start) + head + newRuns + tail + xml.slice(tok.end);
  return { xml: newXml, ok: true, message: "вставлено перед якорем", orderKey: tok.start };
}

/**
 * Найти абзац <w:p>, содержащий заданную литеральную фразу, и вернуть его
 * границы + позицию. Используется для замены/локации пунктов и терминов.
 */
export function findParagraphContaining(
  xml: string,
  phrase: string,
  from = 0,
): { start: number; end: number; inner: string } | null {
  const needle = phrase.replace(/\s+/g, "");
  const pRe = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;
  pRe.lastIndex = from;
  let m: RegExpExecArray | null;
  while ((m = pRe.exec(xml)) !== null) {
    const texts: string[] = [];
    let wt: RegExpExecArray | null;
    WT_RE.lastIndex = 0;
    while ((wt = WT_RE.exec(m[0])) !== null) texts.push(decodeXml(wt[2]));
    const flat = texts.join("").replace(/\s+/g, "");
    if (flat.includes(needle)) {
      return { start: m.index, end: m.index + m[0].length, inner: m[0] };
    }
  }
  return null;
}

/**
 * Найти абзац, чей видимый текст НАЧИНАЕТСЯ с заданного префикса
 * (без учёта пробелов и ведущей автонумерации). Точнее, чем «содержит»:
 * пункты вроде 2.44/1.2 нумеруются автоматически, поэтому опираемся на
 * начало текста пункта (термин, первые слова редакции).
 */
export function findParagraphStartingWith(
  xml: string,
  prefix: string,
): { start: number; end: number; inner: string } | null {
  const needle = prefix.replace(/\s+/g, "");
  if (!needle) return null;
  const pRe = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;
  let m: RegExpExecArray | null;
  let fallback: { start: number; end: number; inner: string } | null = null;
  while ((m = pRe.exec(xml)) !== null) {
    const texts: string[] = [];
    let wt: RegExpExecArray | null;
    WT_RE.lastIndex = 0;
    while ((wt = WT_RE.exec(m[0])) !== null) texts.push(decodeXml(wt[2]));
    const flat = texts.join("").replace(/\s+/g, "");
    if (flat.startsWith(needle)) {
      return { start: m.index, end: m.index + m[0].length, inner: m[0] };
    }
    if (!fallback && flat.includes(needle)) {
      fallback = { start: m.index, end: m.index + m[0].length, inner: m[0] };
    }
  }
  return fallback;
}

/** Собрать весь видимый текст абзаца. */
export function paragraphText(pXml: string): string {
  const texts: string[] = [];
  let wt: RegExpExecArray | null;
  WT_RE.lastIndex = 0;
  while ((wt = WT_RE.exec(pXml)) !== null) texts.push(decodeXml(wt[2]));
  return texts.join("");
}

/**
 * Видимый текст ячейки таблицы.
 *
 * Раны ВНУТРИ абзаца склеиваются без разделителя: число «28», разбитое Word на
 * раны «2» и «8», иначе превратилось бы в «2 8». А границы АБЗАЦЕВ дают
 * пробел: в ячейке приложения на отдельных абзацах стоят название, сайт и
 * приложение партнёра, и без разделителя получалось «ПАО Сбербанкhttps://…».
 * Именно из-за этого расхождения строка из документа «Изменения» не
 * опознавалась как уже присутствующая в таблице и добавлялась второй раз.
 */
export function tableCellText(tcXml: string): string {
  const paras = tcXml.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g);
  const src = paras && paras.length ? paras : [tcXml];
  return src
    .map((p) => paragraphText(p))
    .filter((t) => t.trim() !== "")
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}
