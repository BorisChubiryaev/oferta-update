// Структурная «выжимка» Оферты для ИИ.
//
// Модель не должна получать Оферту целиком: это сотни страниц, и дело не только
// в лимите контекста. Чем больше юридического текста видит модель, тем выше
// соблазн «улучшить» формулировку — а править текст Оферты вправе только
// детерминированный движок. Поэтому наружу отдаётся КАРТА документа: какие
// пункты существуют, как называются разделы, сколько сносок, какие приложения.
// Этого достаточно, чтобы модель указала координаты правки, и недостаточно,
// чтобы она переписала договор.
import { indexNumberedParagraphs } from "./numbering";
import { indexFootnotes } from "./offer-index";
import type { DocxParts } from "./docx";

export interface OfferDigest {
  /** Номера пунктов Оферты в порядке следования («1.1», «2.44», «7.6» …). */
  points: string[];
  /** Заголовки разделов («5. СЕРВИСЫ, ПРЕДОСТАВЛЯЕМЫЕ …»). */
  sections: string[];
  /** Сколько сносок в Оферте (видимая нумерация 1..N). */
  footnoteCount: number;
  /** Номера приложений, найденные в тексте («1», «2» …). */
  appendices: string[];
  /** Термины раздела «Термины и определения»: номер пункта -> термин. */
  terms: { point: string; term: string }[];
  /** Первые слова преамбулы — чтобы модель узнала её, если правка без номера. */
  preambleHead: string;
}

/**
 * Заголовок раздела Оферты.
 *
 * Номер раздела рисует нумерация Word, в тексте абзаца его нет, — поэтому
 * опознаём заголовок по виду: прописные буквы, минимум два слова. Одного
 * «всё прописными» мало: в таблице Приложения № 1 прописными набраны
 * наименования компаний («ООО «МАРКЕТПЛЕЙС»»), и они шли бы в список
 * разделов. Абзацы внутри таблиц отсекаются отдельно (см. tableSpans).
 */
function isSectionHeading(text: string): boolean {
  const t = text.trim();
  if (t.length < 12 || t.length > 160) return false;
  if (t !== t.toUpperCase()) return false;
  if (!/[А-ЯЁ]/.test(t)) return false;
  return t.split(/\s+/).length >= 2;
}

const TBL_RE = /<w:tbl>[\s\S]*?<\/w:tbl>/g;

/** Границы таблиц в document.xml — чтобы не принять ячейку за заголовок. */
function tableSpans(documentXml: string): { from: number; to: number }[] {
  const spans: { from: number; to: number }[] = [];
  let m: RegExpExecArray | null;
  TBL_RE.lastIndex = 0;
  while ((m = TBL_RE.exec(documentXml)) !== null) {
    spans.push({ from: m.index, to: m.index + m[0].length });
  }
  return spans;
}

/**
 * Термин пункта из раздела «Термины»: пункт начинается с термина в кавычках
 * либо с термина до тире. Берём только начало — определение целиком модели не
 * нужно и лишь раздувает запрос.
 */
function termOf(text: string): string | null {
  const q = text.match(/^«([^»]{2,60})»/);
  if (q) return q[1];
  const dash = text.match(/^([А-ЯЁ][^–—-]{2,60}?)\s+[–—-]\s+/);
  return dash ? dash[1].trim() : null;
}

export function buildOfferDigest(offer: DocxParts): OfferDigest {
  const index = indexNumberedParagraphs(offer.document, offer.numbering, offer.styles);
  const spans = tableSpans(offer.document);
  const inTable = (pos: number) => spans.some((s) => pos >= s.from && pos < s.to);
  const points: string[] = [];
  const sections: string[] = [];
  const terms: { point: string; term: string }[] = [];
  const appendices = new Set<string>();

  for (const p of index) {
    if (p.number) {
      points.push(p.number);
      // Раздел «Термины» в Оферте — второй; собирать термины по всему документу
      // незачем, но и жёстко привязываться к номеру раздела нельзя: нумерация
      // разделов от редакции к редакции сдвигается. Признак термина — сам вид
      // абзаца (кавычки или тире в начале), а не его координата.
      const t = termOf(p.text);
      if (t) terms.push({ point: p.number, term: t });
    } else if (!inTable(p.start) && isSectionHeading(p.text)) {
      sections.push(p.text.trim());
    }
    const app = p.text.match(/Приложени[ея]\s*№?\s*(\d+)/i);
    if (app) appendices.add(app[1]);
  }

  const firstText = index.find((p) => p.text.length > 40 && !inTable(p.start));
  return {
    points,
    sections,
    footnoteCount: indexFootnotes(offer.document).displayToId.size,
    appendices: [...appendices].sort((a, b) => Number(a) - Number(b)),
    terms,
    preambleHead: firstText ? firstText.text.slice(0, 400) : "",
  };
}

/**
 * Выжимка в компактном виде для промпта.
 *
 * Номера пунктов сворачиваются в диапазоны: перечислять «7.1, 7.2, … 7.48»
 * поштучно — это тысячи токенов на каждый запрос при нулевой пользе, модели
 * достаточно знать, что пункты 7.1–7.48 существуют.
 */
export function digestToPrompt(d: OfferDigest): string {
  const lines: string[] = [];
  lines.push("РАЗДЕЛЫ ОФЕРТЫ:");
  lines.push(d.sections.length ? d.sections.map((s) => `  ${s}`).join("\n") : "  (не распознаны)");
  lines.push("");
  lines.push(`ПУНКТЫ (${d.points.length} шт.): ${collapseRanges(d.points)}`);
  lines.push(`СНОСКИ: 1–${d.footnoteCount}`);
  lines.push(`ПРИЛОЖЕНИЯ: ${d.appendices.length ? d.appendices.join(", ") : "(не найдены)"}`);
  if (d.terms.length) {
    lines.push("");
    lines.push("ТЕРМИНЫ:");
    for (const t of d.terms.slice(0, 120)) lines.push(`  п. ${t.point} — «${t.term}»`);
  }
  return lines.join("\n");
}

/** «1.1, 1.2, 1.3, 2.1» -> «1.1–1.3, 2.1». */
export function collapseRanges(numbers: string[]): string {
  const out: string[] = [];
  let runStart: string | null = null;
  let prev: string | null = null;
  const tail = (s: string) => {
    const parts = s.split(".");
    const n = Number(parts[parts.length - 1]);
    return Number.isFinite(n) ? n : NaN;
  };
  const head = (s: string) => s.slice(0, s.lastIndexOf("."));
  const flush = () => {
    if (runStart == null || prev == null) return;
    out.push(runStart === prev ? runStart : `${runStart}–${prev}`);
    runStart = null;
  };
  for (const n of numbers) {
    const consecutive =
      prev != null && head(prev) === head(n) && tail(n) === tail(prev) + 1;
    if (!consecutive) {
      flush();
      runStart = n;
    }
    prev = n;
  }
  flush();
  return out.join(", ");
}
