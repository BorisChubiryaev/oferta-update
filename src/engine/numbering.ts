// Восстановление автонумерации Word.
//
// Номера пунктов (2.44, 7.6 …) в document.xml как текст ОТСУТСТВУЮТ — их рисует
// Word по numbering.xml. Чтобы находить пункт «по номеру» (а не по новому
// тексту), нумерацию приходится воспроизводить. Три вещи, без которых
// реконструкция расходится с тем, что видит человек в Word:
//
//  • нумерация может задаваться СТИЛЕМ абзаца, а не самим абзацем: заголовки
//    разделов Оферты не имеют <w:numPr>, зато их стиль ссылается на numId 3
//    уровня 0 — именно они переключают счётчик разделов, и без них все пункты
//    документа выглядят как «1.x»;
//  • номер собирается по шаблону <w:lvlText> («%1.%2.», но встречается и
//    «5.3.%1.» с зашитой частью номера), а не простым склеиванием уровней;
//  • формат уровня <w:numFmt> бывает не только decimal — буквы и римские
//    цифры дают совсем другой видимый номер, а маркеры номера не дают вовсе.
import { decodeXml } from "./ooxml";

const WT_RE = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
const P_RE = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;

interface Level {
  start: number;
  fmt: string;
  text: string | null; // шаблон w:lvlText
}

interface NumDefs {
  /** numId -> уровни (с учётом переопределений конкретного экземпляра). */
  levels: Map<string, Map<number, Level>>;
  /** styleId -> {numId, ilvl} для нумерации, заданной стилем абзаца. */
  styleNum: Map<string, { numId: string; ilvl: number }>;
}

function parseLevels(xml: string): Map<number, Level> {
  const levels = new Map<number, Level>();
  const lvlRe = /<w:lvl\s+w:ilvl="(\d+)"[^>]*>([\s\S]*?)<\/w:lvl>/g;
  let l: RegExpExecArray | null;
  while ((l = lvlRe.exec(xml)) !== null) {
    const body = l[2];
    const st = body.match(/<w:start\s+w:val="(-?\d+)"/);
    const fmt = body.match(/<w:numFmt\s+w:val="([^"]+)"/);
    const text = body.match(/<w:lvlText\s+w:val="([^"]*)"/);
    levels.set(parseInt(l[1], 10), {
      start: st ? parseInt(st[1], 10) : 1,
      fmt: fmt ? fmt[1] : "decimal",
      text: text ? text[1] : null,
    });
  }
  return levels;
}

export function parseNumbering(
  numberingXml: string | null,
  stylesXml: string | null = null,
): NumDefs {
  const levels = new Map<string, Map<number, Level>>();
  const styleNum = new Map<string, { numId: string; ilvl: number }>();
  if (!numberingXml) return { levels, styleNum };

  const abstracts = new Map<string, Map<number, Level>>();
  const absRe = /<w:abstractNum\s+w:abstractNumId="(\d+)"[\s\S]*?<\/w:abstractNum>/g;
  let a: RegExpExecArray | null;
  while ((a = absRe.exec(numberingXml)) !== null) {
    abstracts.set(a[1], parseLevels(a[0]));
  }

  const numRe = /<w:num\s+w:numId="(\d+)"[^>]*>([\s\S]*?)<\/w:num>/g;
  let n: RegExpExecArray | null;
  while ((n = numRe.exec(numberingXml)) !== null) {
    const numId = n[1];
    const absId = n[2].match(/<w:abstractNumId\s+w:val="(\d+)"/);
    const base = absId ? abstracts.get(absId[1]) : undefined;
    const own = new Map<number, Level>(base ?? []);
    // Экземпляр списка может переопределить начало уровня.
    const ovRe = /<w:lvlOverride\s+w:ilvl="(\d+)"[^>]*>([\s\S]*?)<\/w:lvlOverride>/g;
    let o: RegExpExecArray | null;
    while ((o = ovRe.exec(n[2])) !== null) {
      const ilvl = parseInt(o[1], 10);
      const startOv = o[2].match(/<w:startOverride\s+w:val="(-?\d+)"/);
      const inner = parseLevels(o[2]).get(ilvl);
      const cur = own.get(ilvl) ?? { start: 1, fmt: "decimal", text: null };
      own.set(ilvl, {
        start: startOv ? parseInt(startOv[1], 10) : (inner?.start ?? cur.start),
        fmt: inner?.fmt ?? cur.fmt,
        text: inner?.text ?? cur.text,
      });
    }
    levels.set(numId, own);
  }

  if (stylesXml) {
    const stRe = /<w:style\s[^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g;
    let s: RegExpExecArray | null;
    while ((s = stRe.exec(stylesXml)) !== null) {
      const numPr = s[2].match(/<w:numPr>[\s\S]*?<\/w:numPr>/);
      if (!numPr) continue;
      const numId = numPr[0].match(/<w:numId\s+w:val="(\d+)"/);
      if (!numId || numId[1] === "0") continue;
      const ilvl = numPr[0].match(/<w:ilvl\s+w:val="(\d+)"/);
      styleNum.set(s[1], { numId: numId[1], ilvl: ilvl ? parseInt(ilvl[1], 10) : 0 });
    }
  }
  return { levels, styleNum };
}

const LETTERS = "abcdefghijklmnopqrstuvwxyz";
const ROMAN: [number, string][] = [
  [1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"],
  [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"],
];

function formatCounter(value: number, fmt: string): string {
  if (fmt === "lowerLetter" || fmt === "upperLetter") {
    let n = Math.max(1, value) - 1;
    let out = "";
    do {
      out = LETTERS[n % 26] + out;
      n = Math.floor(n / 26) - 1;
    } while (n >= 0);
    return fmt === "upperLetter" ? out.toUpperCase() : out;
  }
  if (fmt === "lowerRoman" || fmt === "upperRoman") {
    let n = Math.max(1, value);
    let out = "";
    for (const [v, sym] of ROMAN) {
      while (n >= v) {
        out += sym;
        n -= v;
      }
    }
    return fmt === "upperRoman" ? out.toUpperCase() : out;
  }
  return String(value);
}

/** Собрать видимый номер уровня по шаблону w:lvlText. */
function renderNumber(
  counters: number[],
  ilvl: number,
  levels: Map<number, Level>,
): string | null {
  const level = levels.get(ilvl);
  if (!level) return null;
  if (level.fmt === "bullet" || level.fmt === "none") return null;
  const template =
    level.text ||
    Array.from({ length: ilvl + 1 }, (_, i) => `%${i + 1}.`).join("");
  const rendered = template.replace(/%(\d)/g, (_m, d: string) => {
    const i = parseInt(d, 10) - 1;
    const lf = levels.get(i)?.fmt ?? "decimal";
    // Уровень с numFmt="none" Word не печатает вовсе. Без этого шаблон
    // «%15.%2.» (то есть %1 + литерал «5.» + %2) читается как «15.N.»
    // вместо «5.N.» — и пункт 5.5 перестаёт находиться по номеру.
    if (lf === "none" || lf === "bullet") return "";
    const value = counters[i] ?? levels.get(i)?.start ?? 1;
    return formatCounter(value, lf);
  });
  const clean = rendered.trim();
  // Шаблон без единого счётчика (маркер-символ) номером не является.
  return /\d|[a-zA-Zа-яА-Я]/.test(clean) ? clean : null;
}

/** Ключ «зрительной» последовательности уровня — либо null.
 *
 *  Бывает, что соседние пункты одного раздела заведены РАЗНЫМИ списками
 *  (5.1 — один numId, 5.2–5.6 — другой). Каждый список считает с единицы, и
 *  реконструкция даёт два пункта «5.1». Но если номер уровня зависит только
 *  от собственного счётчика (все ссылки на верхние уровни — это уровни с
 *  numFmt="none", то есть константа) и в шаблоне зашит номер раздела, то
 *  списки с одинаковым шаблоном человек видит как одну сквозную нумерацию.
 *  Такие уровни считаем общим счётчиком.
 *
 *  Требование «в шаблоне есть зашитая цифра» важно: без него один ключ
 *  получили бы все обычные списки вида «%1.» — включая заголовки разделов.
 */
function sharedSequenceKey(ilvl: number, levels: Map<number, Level>): string | null {
  const level = levels.get(ilvl);
  if (!level?.text) return null;
  const literal = level.text.replace(/%\d/g, "");
  if (!/\d/.test(literal)) return null;
  let usesOwn = false;
  for (const m of level.text.matchAll(/%(\d)/g)) {
    const i = parseInt(m[1], 10) - 1;
    if (i === ilvl) {
      usesOwn = true;
      continue;
    }
    const f = levels.get(i)?.fmt;
    if (f !== "none" && f !== "bullet") return null;
  }
  return usesOwn ? `${ilvl}|${level.text}` : null;
}

export interface NumberedPara {
  start: number;
  end: number;
  inner: string;
  text: string;
  number: string | null; // «2.44», «1.2» … либо null, если абзац не нумерован
  /** Список и уровень, из которых абзац берёт номер (нужно, чтобы понять,
   *  где кончается пункт: подпункты другого списка — часть того же пункта). */
  numId: string | null;
  ilvl: number | null;
}

function paraText(pXml: string): string {
  const parts: string[] = [];
  let m: RegExpExecArray | null;
  WT_RE.lastIndex = 0;
  while ((m = WT_RE.exec(pXml)) !== null) parts.push(decodeXml(m[1]));
  return parts.join("").replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim();
}

/** Нумерация абзаца: собственная (<w:numPr>) либо унаследованная от стиля. */
function paragraphNumbering(
  inner: string,
  defs: NumDefs,
): { numId: string; ilvl: number } | null {
  const pPr = inner.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
  if (!pPr) return null;
  const numPr = pPr[0].match(/<w:numPr>[\s\S]*?<\/w:numPr>/);
  if (numPr) {
    const numId = numPr[0].match(/<w:numId\s+w:val="(\d+)"/);
    if (numId && numId[1] !== "0") {
      const ilvl = numPr[0].match(/<w:ilvl\s+w:val="(\d+)"/);
      return { numId: numId[1], ilvl: ilvl ? parseInt(ilvl[1], 10) : 0 };
    }
    // numId=0 — явное отключение нумерации у абзаца.
    if (numId) return null;
  }
  const style = pPr[0].match(/<w:pStyle\s+w:val="([^"]+)"/);
  if (style) {
    const viaStyle = defs.styleNum.get(style[1]);
    if (viaStyle) return viaStyle;
  }
  return null;
}

/** Проиндексировать все абзацы документа с вычисленными номерами. */
export function indexNumberedParagraphs(
  documentXml: string,
  numberingXml: string | null,
  stylesXml: string | null = null,
): NumberedPara[] {
  const defs = parseNumbering(numberingXml, stylesXml);
  const counters = new Map<string, number[]>(); // numId -> счётчики по уровням
  const shared = new Map<string, number>(); // сквозные счётчики (см. sharedSequenceKey)
  const out: NumberedPara[] = [];

  let m: RegExpExecArray | null;
  P_RE.lastIndex = 0;
  while ((m = P_RE.exec(documentXml)) !== null) {
    const inner = m[0];
    let number: string | null = null;
    const np = paragraphNumbering(inner, defs);
    if (np) {
      const levels = defs.levels.get(np.numId) ?? new Map<number, Level>();
      const startOf = (lv: number) => levels.get(lv)?.start ?? 1;
      const c = counters.get(np.numId) ?? [];
      const seqKey = sharedSequenceKey(np.ilvl, levels);
      if (seqKey) {
        const prev = shared.get(seqKey);
        const value = prev == null ? startOf(np.ilvl) : prev + 1;
        shared.set(seqKey, value);
        c[np.ilvl] = value;
      } else if (c[np.ilvl] == null) c[np.ilvl] = startOf(np.ilvl);
      else c[np.ilvl] = c[np.ilvl] + 1;
      // Вложенные уровни начинают счёт заново.
      for (let k = np.ilvl + 1; k < c.length; k++) c[k] = undefined as unknown as number;
      counters.set(np.numId, c);
      number = renderNumber(c, np.ilvl, levels);
    }
    out.push({
      start: m.index,
      end: m.index + inner.length,
      inner,
      text: paraText(inner),
      number,
      numId: np ? np.numId : null,
      ilvl: np ? np.ilvl : null,
    });
  }
  return out;
}

/** Нормализовать номер к виду «1.2» (убрать хвостовые точки/пробелы). */
export function normNumber(s: string): string {
  return s.trim().replace(/[.\s)]+$/, "");
}

/** Найти абзац по номеру (напр. «1.2»), начиная с позиции fromOffset. */
export function findByNumber(
  index: NumberedPara[],
  number: string,
  fromOffset = 0,
): NumberedPara | null {
  const want = normNumber(number);
  for (const p of index) {
    if (p.start < fromOffset) continue;
    if (p.number && normNumber(p.number) === want) return p;
  }
  return null;
}
