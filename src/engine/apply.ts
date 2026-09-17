// Детерминированное применение структурных операций к Оферте.
// Текст берётся ТОЛЬКО из операций (payload/rows) — движок не «сочиняет».
import type { DocxParts } from "./docx";
import { saveDocx } from "./docx";
import {
  prepareComments,
  addComment,
  commentRangeStart,
  commentRangeEnd,
  type CommentsState,
} from "./comments";
import { indexFootnotes, findFootnoteById, allFootnotes } from "./offer-index";
import {
  insertAfterAnchor,
  insertBeforeAnchor,
  replacePhraseRuns,
  phraseOccurrenceAfter,
  countPhrase,
  paragraphText,
  escapeXml,
} from "./ooxml";
import {
  locateReplaceParagraph,
  locatePointInsertion,
  locatePointSpan,
  locatePointBlock,
  locateByTextPrefix,
  locatePreamble,
  appendixOffset,
} from "./locate";
import type { ParaSpan } from "./locate";
import { indexNumberedParagraphs } from "./numbering";
import { findAppendixTable, appendixHeadingOffset, replaceRows, buildRow } from "./tables";
import {
  sortTableAlphabetically,
  alphabeticalPosition,
  isAlphabeticallyOrdered,
  setRowNumber,
  parseRows,
  findExistingRow,
} from "./alpha-sort";
import {
  maxFootnoteId,
  footnoteRefRunRpr,
  buildFootnoteReferenceRun,
  buildFootnoteElement,
  appendFootnoteElement,
} from "./footnote-add";
import {
  renderInsertRuns,
  renderDeleteRuns,
  renderOldPrefix,
  renderOldInline,
  resetInsCounter,
} from "./render";
import type { ApplyResult, BuildOptions, Operation } from "./types";

/**
 * Таблица, к которой относится операция.
 *
 * Если правка называет пункт («дополнить таблицу в п.4 Приложения № 2»), ищем
 * таблицу НИЖЕ этого пункта: в приложении таблиц несколько, и без такого
 * уточнения строка уходила в первую попавшуюся — например, строка перечня
 * Посредников (4 колонки) приписывалась в конец таблицы п.3 (6 колонок).
 * Номера пунктов внутри приложения проставлены автонумерацией, поэтому пункт
 * ищем тем же движком нумерации, что и обычные пункты Оферты.
 */
function tableForOp(op: Operation, state: ApplyState) {
  const appendix = op.target.kind === "appendix_table" ? op.target.appendix : "2";
  const point = op.target.kind === "appendix_table" ? op.target.point : undefined;
  if (point) {
    const head = appendixHeadingOffset(state.document, appendix);
    const span = locatePointSpan(state.document, state.numbering, point, state.styles, head);
    if (span) return findAppendixTable(state.document, appendix, span.end);
  }
  return findAppendixTable(state.document, appendix);
}

/** Как в сообщении описать судьбу прежнего текста при замене. */
function oldTextNote(opts: BuildOptions): string {
  return opts.showOld === false && (opts.highlightMode ?? "color") === "color"
    ? "прежняя редакция скрыта"
    : "прежняя редакция показана зачёркнутой";
}

/** Заменить содержимое сноски (все раны) на новый текст, сохранив маркер. */
function replaceFootnoteBody(inner: string, text: string, opts: BuildOptions): string {
  // Оставляем служебный первый ран со ссылкой-номером, если он есть; иначе
  // просто заменяем все текстовые раны одним выделенным.
  const runs = renderInsertRuns(text, opts);
  const ref = inner.match(/<w:r\b[^>]*>[\s\S]*?<w:footnoteRef\s*\/>[\s\S]*?<\/w:r>/);
  const firstP = inner.match(/<w:p\b[^>]*>/);
  const open = firstP ? firstP[0] : "<w:p>";
  const pPr = inner.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
  // Старый текст сноски показываем зачёркнутым перед новым — тот же принцип
  // «видно, как было», что и для обычных пунктов.
  const oldPlain = currentText(inner);
  const oldRuns = renderOldPrefix(oldPlain, opts);
  return `${open}${pPr ? pPr[0] : ""}${ref ? ref[0] : ""}${oldRuns}${runs}</w:p>`;
}

/**
 * Свойства абзаца с ЯВНО отключённой нумерацией.
 *
 * Просто убрать <w:numPr> недостаточно: нумерация может приходить из стиля
 * абзаца, и абзац без собственного numPr молча становится новым пунктом (а со
 * стилем заголовка — новым разделом, сдвигая нумерацию всего документа).
 * Отключение в OOXML — это numId="0", его и ставим.
 */
function withoutNumbering(pPr: string): string {
  const cleaned = pPr.replace(/<w:numPr>[\s\S]*?<\/w:numPr>/g, "");
  const off = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>';
  if (!cleaned) return `<w:pPr>${off}</w:pPr>`;
  // По схеме numPr идёт сразу после pStyle.
  const style = cleaned.match(/<w:pStyle\b[^>]*\/>/);
  if (style) return cleaned.replace(style[0], style[0] + off);
  return cleaned.replace("<w:pPr>", "<w:pPr>" + off);
}

/** Извлечь <w:pPr>…</w:pPr> из начала абзаца (стиль/нумерация сохраняются). */
function extractPPr(pXml: string): string {
  const m = pXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
  return m ? m[0] : "";
}

/** Заменить все раны абзаца на новые, сохранив pPr. */
function replaceParagraphRuns(pXml: string, newRuns: string): string {
  const pPr = extractPPr(pXml);
  const openMatch = pXml.match(/^<w:p(?:\s[^>]*)?>/);
  const open = openMatch ? openMatch[0] : "<w:p>";
  return `${open}${pPr}${newRuns}</w:p>`;
}

/** Убрать ведущий номер пункта («2.44.», «7.6 ») — он даётся автонумерацией. */
/** Совпадение текстов «по существу»: без номера, кавычек-ёлочек и пробелов. */
/**
 * «Актуальный» видимый текст абзаца/сноски: без содержимого, помеченного как
 * удалённое в ПРЕДЫДУЩЕМ раунде правок — зачёркнутого рана (цветной режим)
 * или <w:del> (режим рецензирования, его paragraphText и так не берёт, т.к.
 * там <w:delText>, а не <w:t>). Без этой фильтрации повторный прогон на уже
 * отредактированной Оферте склеивает «было»+«стало» в одну строку, сравнение
 * с новой редакцией никогда не совпадает, и старый текст задваивается на
 * каждом следующем раунде правок.
 */
function currentText(innerXml: string): string {
  let out = "";
  const re = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(innerXml)) !== null) {
    if (/<w:strike\s*\/>/.test(m[0])) continue;
    out += m[0];
  }
  return paragraphText(out).replace(/\s+/g, " ").trim();
}

function sameText(a: string, b: string): boolean {
  const norm = (t: string) =>
    stripLeadingNumber(t)
      .replace(/[«»"']/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  const x = norm(a);
  const y = norm(b);
  if (!x || !y) return false;
  return x === y || x.startsWith(y) || y.startsWith(x);
}

/**
 * Снять номер пункта в начале новой редакции.
 *
 * Номер абзацу присваивает Word по автонумерации, поэтому в тексте он лишний —
 * иначе в Оферте получается «9.6 9.6 Клиент проинформирован…». Сокращение
 * «п.»/«п.п.» перед номером встречается в исходниках наравне с голым номером
 * («…изложив в следующей редакции: «п.9.6 Клиент проинформирован…»»), и без
 * его учёта номер уезжал в текст пункта.
 */
function stripLeadingNumber(text: string): string {
  return text.replace(/^\s*(?:п\.?\s*п\.?|п\.?)?\s*\d+(?:\.\d+)*\.?\s*/i, "");
}

/** Снять внешние кавычки-«ёлочки», если инструкция дала payload целиком в них. */
function stripOuterQuotes(text: string): string {
  const t = text.trim();
  if (t.startsWith("«") && t.endsWith("»")) return t.slice(1, -1);
  return t;
}

/** Собрать выделенные раны абзаца; для термина ведущее слово — жирным. */
function buildParagraphRuns(body: string, termLike: boolean, opts: BuildOptions): string {
  if (termLike) {
    const dashIdx = body.search(/\s[–—-]\s/);
    if (dashIdx > 0) {
      const term = body.slice(0, dashIdx);
      const rest = body.slice(dashIdx);
      return (
        renderInsertRuns(term, opts).replace("<w:rPr>", "<w:rPr><w:b/>") +
        renderInsertRuns(rest, opts)
      );
    }
  }
  return renderInsertRuns(body, opts);
}

/**
 * Варианты якорной фразы с усечённым окончанием последнего слова.
 *
 * В документах «Изменения» падеж последнего слова нередко расходится с текстом
 * Оферты — «Расчет скоринговой оценка» против «…оценки». Отбрасывать окончание
 * можно только у последнего слова и не больше двух букв, а результат
 * применяется, лишь если он найден ВНУТРИ уже определённого пункта: это
 * поблажка к грамматике, а не поиск похожего места.
 */
function anchorVariants(anchor: string): string[] {
  const trimmed = anchor.trim();
  const m = trimmed.match(/([А-Яа-яЁёA-Za-z]{4,})$/);
  if (!m) return [trimmed];
  const base = trimmed.slice(0, trimmed.length - m[1].length);
  return [trimmed, base + m[1].slice(0, -1), base + m[1].slice(0, -2)];
}

/** Вставка с поблажкой к падежному окончанию якоря. */
function insertNearAnchor(
  xml: string,
  anchor: string,
  runs: string,
  before: boolean,
): { result: InsertLike; usedVariant: string } | null {
  const insertAt = before ? insertBeforeAnchor : insertAfterAnchor;
  for (const variant of anchorVariants(anchor)) {
    const res = insertAt(xml, variant, runs);
    if (res.ok) return { result: res, usedVariant: variant };
  }
  return null;
}

interface InsertLike {
  xml: string;
  ok: boolean;
  message: string;
  orderKey: number;
}

/** Раны со ссылками на сноски — их нельзя терять при замене абзаца. */
function footnoteRefRuns(pXml: string): string {
  const runs = pXml.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g) ?? [];
  return runs.filter((r) => /<w:footnoteReference\b/.test(r)).join("");
}

/** С какого места документа искать цель: пункт приложения — ниже его заголовка. */
function searchFrom(op: Operation, document: string): number {
  return op.target.kind === "appendix_point" || op.target.kind === "appendix_table"
    ? appendixOffset(document, op.target.appendix)
    : 0;
}

/**
 * Видимый номер сноски для операции. Сноску можно задать номером
 * («сноску 4») или пунктом («сноска к пункту 5.3») — во втором случае берём
 * первую ссылку на сноску внутри этого пункта.
 */
function resolveFootnoteNumber(op: Operation, state: ApplyState): number | null {
  if (op.target.kind !== "footnote") return null;
  if (op.target.number > 0) return op.target.number;
  const point = op.target.atPoint;
  if (!point) return null;
  const block = locatePointBlock(state.document, state.numbering, point, state.styles);
  if (!block.length) return null;
  const idx = indexFootnotes(state.document);
  for (const [display, pos] of idx.displayToBodyPos) {
    if (block.some((b) => pos >= b.start && pos < b.end)) return display;
  }
  return null;
}

/** Номера пунктов, в тексте которых встречается фраза. */
function pointsContaining(state: ApplyState, phrase: string): string[] {
  const index = indexNumberedParagraphs(state.document, state.numbering, state.styles);
  const out: string[] = [];
  let current: string | null = null;
  for (const p of index) {
    if (p.number) current = p.number;
    if (current && countPhrase(p.inner, phrase) > 0 && !out.includes(current)) out.push(current);
  }
  return out;
}

/** Номер пункта, на который направлена операция (если он есть). */
function opPoint(op: Operation): string | null {
  if (op.target.kind === "point" || op.target.kind === "appendix_point") return op.target.point;
  if (op.target.kind === "term" && op.target.point) return op.target.point;
  return null;
}

/**
 * Порядковый номер в человеческом виде: 1 — первый, −1 — последний.
 * Отрицательные значения считаются с конца, как в срезах.
 */
function pickIndex(wanted: number, length: number): number {
  return wanted < 0 ? length + wanted : wanted - 1;
}

/** Порядковое слово с согласованием: абзац мужского рода, предложение — среднего. */
function ordinalWord(n: number, gender: "m" | "n" = "n"): string {
  const forms = {
    n: ["первое", "второе", "третье", "четвёртое", "пятое", "шестое"],
    m: ["первый", "второй", "третий", "четвёртый", "пятый", "шестой"],
  }[gender];
  if (n < 0) return gender === "m" ? "последний" : "последнее";
  return forms[n - 1] ?? `${n}-й`;
}

/**
 * Разбить текст пункта на предложения.
 *
 * Точка в «п. 5.3» или «т.д.» — не конец предложения, поэтому границей
 * считаем точку, за которой идёт пробел и заглавная буква либо «ёлочка».
 */
function sentences(text: string): string[] {
  return text
    .split(/(?<=[.!?])\s+(?=[«“(]?[А-ЯЁA-Z])/)
    .map((x) => x.trim())
    .filter(Boolean);
}

/** Заменить фрагмент документа на новый XML абзаца. */
function spliceSpan(document: string, span: ParaSpan, rebuilt: string): string {
  return document.slice(0, span.start) + rebuilt + document.slice(span.end);
}

export interface ApplyState {
  document: string;
  footnotes: string | null;
  numbering: string | null;
  styles: string | null;
  /**
   * Комментарии Word для правок, которые движок не вносит сам. Отсутствуют при
   * предпросмотре: он гоняет операции по копии состояния только ради текста и
   * порядка, а части пакета не собирает.
   */
  comments?: CommentsState;
}

export function applyOneOp(
  op: Operation,
  state: ApplyState,
  opts: BuildOptions,
): ApplyResult {
  const fail = (message: string): ApplyResult => ({
    operationId: op.id,
    ok: false,
    message,
    orderKey: Number.MAX_SAFE_INTEGER,
  });

  // ── Вставка относительно якоря ─────────────────────────────────────
  if (op.type === "insert_after" || op.type === "insert_before") {
    if (!op.anchor || op.payload === undefined) return fail("нет якоря/текста");
    const runs = renderInsertRuns(op.payload, opts);

    if (op.target.kind === "footnote") {
      if (!state.footnotes) return fail("в документе нет сносок");
      const idx = indexFootnotes(state.document);
      const id = idx.displayToId.get(op.target.number);
      let note = "";
      const fn = id !== undefined ? findFootnoteById(state.footnotes, id) : null;
      // Пробуем вставить в сноску по номеру.
      if (fn) {
        const res = insertAfterAnchor(fn.inner, op.anchor, runs);
        if (res.ok) {
          state.footnotes =
            state.footnotes.slice(0, fn.start) +
            state.footnotes.slice(fn.start).replace(fn.inner, res.xml);
          return {
            operationId: op.id,
            ok: true,
            message: `сноска № ${op.target.number}: вставлено`,
            orderKey: idx.displayToBodyPos.get(op.target.number) ?? fn.start,
          };
        }
      }
      // Номер не совпал (нумерация «поехала») — ищем сноску по якорю. Но
      // только если якорь встречается РОВНО В ОДНОЙ сноске: иначе две правки
      // с одинаковым якорем молча уедут в одну и ту же сноску.
      const blocks = allFootnotes(state.footnotes);
      const matching = blocks.filter((b) => countPhrase(b.inner, op.anchor!) > 0);
      if (matching.length === 1) {
        const b = matching[0];
        const res = insertAfterAnchor(b.inner, op.anchor, runs);
        if (res.ok) {
          state.footnotes =
            state.footnotes.slice(0, b.start) +
            state.footnotes.slice(b.start).replace(b.inner, res.xml);
          note = ` (номер не совпал — найдено по содержимому, сноска id=${b.id})`;
          return {
            operationId: op.id,
            ok: true,
            message: `сноска № ${op.target.number}: вставлено${note}`,
            orderKey: b.start,
          };
        }
      }
      if (matching.length > 1) {
        // Номера кандидатов — это не украшение отчёта: без них оператору
        // пришлось бы искать нужную сноску вручную по всему документу.
        const fnIdx = indexFootnotes(state.document);
        const idToDisplay = new Map<number, number>();
        fnIdx.displayToId.forEach((fid, display) => idToDisplay.set(fid, display));
        const numbers = matching
          .map((b) => idToDisplay.get(b.id))
          .filter((n): n is number => n !== undefined)
          .sort((a, b) => a - b);
        return fail(
          `сноска № ${op.target.number}: в ней якоря «${op.anchor.slice(0, 40)}» нет; ` +
            `он есть в сносках № ${numbers.join(", ")} — укажите нужный номер вручную`,
        );
      }
      return fail(`сноска № ${op.target.number}: якорь «${op.anchor}» не найден ни по номеру, ни по содержимому`);
    }

    // Вставка в тело. Ищем якорь ВНУТРИ целевого пункта: фразы вроде
    // «и услуг Банка,» встречаются в Оферте многократно, и поиск по всему
    // документу молча вставил бы все правки в первое попавшееся место.
    const point = opPoint(op);
    if (point) {
      // Ищем по всему блоку пункта: «пункт 7.6» — это и его подпункты, а
      // якорная фраза нередко лежит именно в подпункте.
      const block = locatePointBlock(
        state.document,
        state.numbering,
        point,
        state.styles,
        searchFrom(op, state.document),
      );
      for (const span of block) {
        // Уже внесено (повторный прогон, или правка была в предыдущей
        // редакции) — не дублируем текст. Проверяем именно СВЯЗКУ «якорь +
        // вставка»: короткая вставка вроде предлога «с» встречается в пункте
        // где угодно, и проверка по одному payload'у ложно срабатывала.
        const plain = paragraphText(span.inner).replace(/\s+/g, " ");
        const already = (op.anchor + " " + op.payload).replace(/[«»\s]+/g, " ").trim();
        if (already.length > 4 && plain.replace(/[«»\s]+/g, " ").includes(already)) {
          return {
            operationId: op.id,
            ok: true,
            message: `п. ${point}: текст уже присутствует — правка не требуется`,
            orderKey: span.start,
          };
        }
        const hit = insertNearAnchor(span.inner, op.anchor, runs, op.type === "insert_before");
        if (hit) {
          state.document = spliceSpan(state.document, span, hit.result.xml);
          const loose = hit.usedVariant !== op.anchor.trim();
          return {
            operationId: op.id,
            ok: true,
            message:
              `п. ${point}: вставлено ${op.type === "insert_before" ? "перед" : "после"} слов${op.type === "insert_before" ? "ами" : ""} «${op.anchor.slice(0, 40)}»` +
              (loose ? " (окончание последнего слова в тексте Оферты другое — сверьте место)" : ""),
            orderKey: span.start,
          };
        }
      }
    }
    // В пункте якоря нет (номера пунктов в разных редакциях расходятся).
    // Тогда опираемся на содержимое — но только если оно однозначно.
    const pointFound = point
      ? !!locatePointSpan(state.document, state.numbering, point, state.styles, searchFrom(op, state.document))
      : false;
    const hits = countPhrase(state.document, op.anchor);
    if (hits === 0)
      return fail(
        `якорь «${op.anchor.slice(0, 60)}» не найден${point ? ` ни в п. ${point}, ни в остальном тексте` : ""}`,
      );
    if (hits > 1) {
      // Подсказываем номера пунктов, где якорь есть: чаще всего расхождение —
      // это сдвиг нумерации на один-два пункта, и оператору достаточно
      // поправить номер в карточке.
      const where = pointsContaining(state, op.anchor).slice(0, 6);
      const hint = where.length ? `; он есть в п. ${where.join(", ")}` : "";
      return fail(
        (pointFound
          ? `п. ${point} найден, но якоря «${op.anchor.slice(0, 40)}» в нём нет`
          : `п. ${point ?? "?"} по этому номеру не найден, а якорь «${op.anchor.slice(0, 40)}» встречается ${hits} раз`) +
          hint +
          " — укажите нужный пункт вручную",
      );
    }
    const res = (op.type === "insert_before" ? insertBeforeAnchor : insertAfterAnchor)(
      state.document,
      op.anchor,
      runs,
    );
    if (!res.ok) return fail(res.message);
    state.document = res.xml;
    return {
      operationId: op.id,
      ok: true,
      message: point
        ? `вставлено по содержимому (пункт с номером ${point} не найден, якорь в тексте единственный)`
        : "вставлено в текст",
      orderKey: res.orderKey,
    };
  }

  // ── Замена пункта / термина / пункта приложения ────────────────────
  if (op.type === "replace") {
    if (op.payload === undefined) return fail("нет текста замены");
    const body = stripLeadingNumber(stripOuterQuotes(op.payload));
    // Цель ищем по НОМЕРУ пункта/имени термина и разделу (а не по новому тексту).
    let byPrefix = false;
    let para = locateReplaceParagraph(state.document, state.numbering, op, state.styles);
    if (!para && op.target.kind !== "preamble") {
      // Номер не нашёлся — пробуем опознать пункт по началу его текста.
      para = locateByTextPrefix(
        state.document,
        state.numbering,
        body,
        state.styles,
        opPoint(op) ?? undefined,
      );
      byPrefix = !!para;
    }
    if (!para) {
      let what = "пункт";
      if (op.target.kind === "term") what = `термин «${op.target.term}»`;
      else if (op.target.kind === "appendix_point")
        what = `пункт ${op.target.point} Приложения №${op.target.appendix}`;
      else if (op.target.kind === "point") what = `пункт ${op.target.point}`;
      return fail(`${what} не найден в документе — проверьте, что правка применима к этой редакции`);
    }

    const oldPlain = currentText(para.inner);
    if (sameText(oldPlain, body)) {
      return {
        operationId: op.id,
        ok: true,
        message: "пункт уже изложен в этой редакции — правка не требуется",
        orderKey: para.start,
      };
    }

    const newRuns = buildParagraphRuns(body, op.target.kind === "term", opts);
    // Ссылки на сноски живут в ранах абзаца: заменив раны целиком, мы бы
    // осиротили сноски и сломали их сквозную нумерацию.
    const keptRefs = footnoteRefRuns(para.inner);
    // Прежний текст показываем зачёркнутым — читатель должен видеть, «как
    // было», а не только итоговую редакцию (иначе правку нельзя проверить
    // без второго открытого окна со старой Офертой).
    const oldRuns = renderOldPrefix(oldPlain, opts);
    // Многоабзацная редакция (преамбула — это перечень Ключевых Компаний, по
    // абзацу на компанию) раскладывается обратно по абзацам: одним абзацем
    // список превратился бы в сплошную простыню текста.
    const extra = body.includes("\n")
      ? body
          .split("\n")
          .slice(1)
          .filter((line) => line.trim())
          .map(
            (line) =>
              `<w:p>${extractPPr(para!.inner)}${buildParagraphRuns(line.trim(), false, opts)}</w:p>`,
          )
          .join("")
      : "";
    const firstLine = body.includes("\n") ? body.slice(0, body.indexOf("\n")).trim() : body;
    const headRuns = body.includes("\n")
      ? buildParagraphRuns(firstLine, op.target.kind === "term", opts)
      : newRuns;
    const rebuilt =
      replaceParagraphRuns(para.inner, oldRuns + headRuns + keptRefs) + extra;
    state.document = spliceSpan(state.document, para, rebuilt);
    return {
      operationId: op.id,
      ok: true,
      message:
        `пункт изложен в новой редакции (${oldTextNote(opts)})` +
        (byPrefix
          ? " (пункт с указанным номером не найден — опознан по началу текста, сверьте место правки)"
          : "") +
        (keptRefs ? " (ссылки на сноски сохранены — проверьте их уместность)" : ""),
      orderKey: para.start,
      oldText: oldPlain,
      newText: body,
    };
  }

  // ── Добавление НОВОГО пункта (нумерация сдвигается автоматически) ───
  if (op.type === "insert_point") {
    if (op.payload === undefined) return fail("нет текста нового пункта");
    const point = opPoint(op);
    if (!point) return fail("не указан номер нового пункта");
    const body = stripLeadingNumber(stripOuterQuotes(op.payload));
    // Правка могла быть уже учтена в загруженной редакции Оферты: пункт с
    // таким номером есть, и текст у него тот же. Это не ошибка размещения —
    // добавлять второй такой же пункт нельзя.
    const existing = locatePointSpan(state.document, state.numbering, point, state.styles);
    if (existing && sameText(paragraphText(existing.inner), body)) {
      return {
        operationId: op.id,
        ok: true,
        message: `п. ${point} уже присутствует в этой редакции — правка не требуется`,
        orderKey: existing.start,
      };
    }
    const loc = locatePointInsertion(state.document, state.numbering, point, state.styles);
    if (!loc) return fail(`не найдено место для нового пункта ${point} (раздел/соседний пункт)`);
    // Новый абзац наследует стиль/нумерацию (numPr) соседнего пункта — тогда
    // Word сам присвоит номер и перенумерует последующие.
    const pPr = extractPPr(loc.span.inner);
    const isTermLike = /\s[–—-]\s/.test(body.slice(0, 80));
    const runs = buildParagraphRuns(body, isTermLike, opts);
    const newPara = `<w:p>${pPr}${runs}</w:p>`;
    const at = loc.mode === "before" ? loc.span.start : loc.span.end;
    state.document = state.document.slice(0, at) + newPara + state.document.slice(at);
    return {
      operationId: op.id,
      ok: true,
      message: `добавлен пункт ${point} (последующие перенумеруются автоматически)`,
      orderKey: at,
    };
  }

  // ── Добавить НОВУЮ сноску (нумерация сносок сдвигается автоматически) ─
  if (op.type === "add_footnote") {
    if (!state.footnotes) return fail("в документе нет блока сносок");
    if (op.payload === undefined || !op.anchor) return fail("нет якоря или текста сноски");
    const id = maxFootnoteId(state.footnotes) + 1;
    const refRun = buildFootnoteReferenceRun(id, footnoteRefRunRpr(state.document));
    // Вставляем ссылку после якоря — по возможности внутри нужного пункта.
    let inserted = false;
    let looseAnchor = false;
    const point =
      op.target.kind === "point" || op.target.kind === "appendix_point" ? op.target.point : undefined;
    if (point) {
      // Ищем якорь по всему блоку пункта — как и при вставке текста: фраза
      // нередко лежит в подпункте. Прежде здесь использовался поиск МЕСТА ДЛЯ
      // НОВОГО пункта, который при отсутствии номера возвращал соседний пункт.
      const block = locatePointBlock(
        state.document,
        state.numbering,
        point,
        state.styles,
        searchFrom(op, state.document),
      );
      for (const span of block) {
        const hit = insertNearAnchor(span.inner, op.anchor, refRun, false);
        if (hit) {
          state.document = spliceSpan(state.document, span, hit.result.xml);
          inserted = true;
          if (hit.usedVariant !== op.anchor.trim()) looseAnchor = true;
          break;
        }
      }
    }
    if (!inserted) {
      const res = insertAfterAnchor(state.document, op.anchor, refRun);
      if (!res.ok)
        return fail(
          `якорь «${op.anchor}» для сноски не найден — возможно, ` +
            "правка рассчитана на другую редакцию Оферты",
        );
      state.document = res.xml;
    }
    state.footnotes = appendFootnoteElement(state.footnotes, buildFootnoteElement(id, op.payload, opts));
    const orderKey = state.document.indexOf(`<w:footnoteReference w:id="${id}"`);
    return {
      operationId: op.id,
      ok: true,
      message:
        "добавлена сноска (последующие сноски перенумеруются автоматически)" +
        (looseAnchor ? "; окончание последнего слова якоря в Оферте другое — сверьте место" : ""),
      orderKey: orderKey >= 0 ? orderKey : 0,
    };
  }

  // ── Замена сноски целиком ──────────────────────────────────────────
  if (op.type === "replace_footnote") {
    if (op.target.kind !== "footnote") return fail("цель не является сноской");
    if (op.payload === undefined) return fail("нет текста замены");
    if (!state.footnotes) return fail("в документе нет сносок");
    const number = resolveFootnoteNumber(op, state) ?? op.target.number;
    const idx = indexFootnotes(state.document);
    const id = idx.displayToId.get(number);
    const fn = id !== undefined ? findFootnoteById(state.footnotes, id) : null;
    if (!fn) return fail(`сноска № ${number} не найдена`);
    const body = stripOuterQuotes(op.payload);
    const oldPlain = currentText(fn.inner);
    if (sameText(oldPlain, body)) {
      return {
        operationId: op.id,
        ok: true,
        message: `сноска № ${number}: уже в этой редакции — правка не требуется`,
        orderKey: idx.displayToBodyPos.get(number) ?? fn.start,
      };
    }
    const rebuilt = replaceFootnoteBody(fn.inner, body, opts);
    state.footnotes =
      state.footnotes.slice(0, fn.start) +
      state.footnotes.slice(fn.start).replace(fn.inner, rebuilt);
    return {
      operationId: op.id,
      ok: true,
      message:
        `сноска № ${number}: изложена в новой редакции (${oldTextNote(opts)})` +
        (op.target.kind === "footnote" && op.target.atPoint
          ? ` (найдена как первая сноска п. ${op.target.atPoint})`
          : ""),
      orderKey: idx.displayToBodyPos.get(number) ?? fn.start,
      oldText: oldPlain,
      newText: body,
    };
  }

  // ── Добавление строк в таблицу приложения ──────────────────────────
  if (op.type === "append_table_rows") {
    if (!op.rows || op.rows.length === 0) return fail("нет строк для добавления");
    const appendix = op.target.kind === "appendix_table" ? op.target.appendix : "2";
    const table = tableForOp(op, state);
    if (!table) return fail(`таблица Приложения №${appendix} не найдена`);
    const nameCol = op.nameColumn ?? 1;
    // Строку, которая в таблице уже есть, добавлять нельзя: правку легко
    // применить к редакции Оферты, где она уже учтена, и тогда в приложении
    // появлялся второй такой же партнёр. Сравниваем по наименованию — как и
    // при алфавитной вставке.
    const add: string[][] = [];
    const skipped: string[] = [];
    for (const cells of op.rows) {
      const name = cells[nameCol] ?? "";
      const dup = name.trim() ? findExistingRow(table.inner, name, nameCol) : null;
      if (dup) skipped.push(`${name.replace(/\s+/g, " ").slice(0, 40)} (уже есть, строка ${dup.number})`);
      else add.push(cells);
    }
    if (!add.length) {
      return {
        operationId: op.id,
        ok: true,
        message: `Приложение №${appendix}: строки уже присутствуют, добавлять нечего — ${skipped.join("; ")}`,
        orderKey: table.start,
      };
    }
    const tblEnd = table.end - "</w:tbl>".length;
    const rowsXml = add.map((r) => buildRow(r, opts)).join("");
    state.document = state.document.slice(0, tblEnd) + rowsXml + state.document.slice(tblEnd);
    return {
      operationId: op.id,
      ok: true,
      message:
        `добавлено строк: ${add.length}` +
        (skipped.length ? `; пропущено как уже имеющиеся: ${skipped.join("; ")}` : ""),
      orderKey: table.start,
    };
  }

  // ── Замена существующих строк таблицы ──────────────────────────────
  if (op.type === "replace_table_rows") {
    if (!op.rows || op.rows.length === 0) return fail("нет данных строк для замены");
    const appendix = op.target.kind === "appendix_table" ? op.target.appendix : "2";
    const table = tableForOp(op, state);
    if (!table) return fail(`таблица Приложения №${appendix} не найдена`);
    const nameCol = op.nameColumn ?? 1;
    const reps = op.rows.map((cells, i) => ({
      number: parseInt((cells[0] || "").trim(), 10) || op.rowNumbers?.[i] || 0,
      cells,
    }));
    const res = replaceRows(table.inner, reps, opts, nameCol);
    state.document =
      state.document.slice(0, table.start) + res.xml + state.document.slice(table.end);

    const moved = res.replaced.filter((r) => r.byName && r.atRow !== r.number);
    const details: string[] = [];
    if (moved.length)
      details.push(
        `сопоставлено по наименованию (номер в правке отличается): ${moved
          .map((r) => `${r.name.slice(0, 24)} №${r.number}→№${r.atRow}`)
          .slice(0, 4)
          .join(", ")}${moved.length > 4 ? ` и ещё ${moved.length - 4}` : ""}`,
      );
    if (res.missing.length)
      details.push(`не найдены: ${res.missing.map((m) => m.name || `№${m.number}`).join(", ")}`);
    if (res.warnings.length) details.push(res.warnings.slice(0, 2).join("; "));

    return {
      operationId: op.id,
      ok: res.replaced.length > 0,
      message:
        `заменено строк: ${res.replaced.length}/${reps.length} в Приложении №${appendix}` +
        (details.length ? `; ${details.join("; ")}` : ""),
      orderKey: table.start,
    };
  }

  // ── Пересортировка таблицы приложения по алфавиту ──────────────────
  if (op.type === "sort_table_alpha") {
    const appendix = op.target.kind === "appendix_table" ? op.target.appendix : "1";
    const table = findAppendixTable(state.document, appendix);
    if (!table) return fail(`таблица Приложения №${appendix} не найдена`);
    const res = sortTableAlphabetically(table.inner, op.nameColumn ?? 1);
    if ("error" in res) return fail(res.error);
    state.document =
      state.document.slice(0, table.start) + res.xml + state.document.slice(table.end);
    const warn = res.warnings.length ? `; ${res.warnings.join("; ")}` : "";
    return {
      operationId: op.id,
      ok: true,
      message:
        res.moves.length === 0
          ? `Приложение №${appendix}: уже в алфавитном порядке, нумерация проверена`
          : `Приложение №${appendix}: отсортировано по алфавиту, перемещено строк: ${res.moves.length}${warn}`,
      orderKey: table.start,
    };
  }

  // ── Вставка строки в таблицу по алфавиту ───────────────────────────
  if (op.type === "insert_table_row_alpha") {
    if (!op.rows || op.rows.length === 0) return fail("нет данных новой строки");
    const appendix = op.target.kind === "appendix_table" ? op.target.appendix : "1";
    const table = findAppendixTable(state.document, appendix);
    if (!table) return fail(`таблица Приложения №${appendix} не найдена`);
    const nameCol = op.nameColumn ?? 1;
    const rowsXml = table.inner.match(/<w:tr\b[\s\S]*?<\/w:tr>/g) ?? [];
    const parsed = parseRows(table.inner);
    const hasHeader = !/^\d+$/.test(parsed[0]?.cells[0] ?? "");
    const headerCount = hasHeader ? 1 : 0;

    const inserted: string[] = [];
    const skipped: string[] = [];
    let out = [...rowsXml];
    // Шапку уже отрезали сами, поэтому говорим об этом явно: иначе строку без
    // номера (а у только что вставленной его ещё нет) примут за шапку.
    const unordered = !isAlphabeticallyOrdered(
      out.slice(headerCount).join(""),
      nameCol,
      false,
    );
    for (const cells of op.rows) {
      const name = cells[nameCol] ?? "";
      if (!name.trim()) continue;
      const body = out.slice(headerCount).join("");
      // Защита от дубликатов: если компания уже в таблице — не добавляем.
      const dup = findExistingRow(body, name, nameCol, false);
      if (dup) {
        skipped.push(`${name} (уже есть, строка ${dup.number})`);
        continue;
      }
      const pos = alphabeticalPosition(body, name, nameCol, false);
      out.splice(headerCount + pos - 1, 0, buildRow(cells, opts));
      inserted.push(`${name} → позиция ${pos}`);
    }
    if (inserted.length === 0) {
      // Все строки уже в таблице — правка учтена в этой редакции Приложения.
      // Это результат, а не сбой: сообщаем, но не помечаем ошибкой.
      return {
        operationId: op.id,
        ok: skipped.length > 0,
        message: skipped.length
          ? `Приложение №${appendix}: уже присутствует, добавлять нечего — ${skipped.join("; ")}`
          : "не удалось определить наименование новой строки",
        orderKey: table.start,
      };
    }
    // Перенумеровываем весь корпус таблицы.
    out = out.map((tr, i) => (i < headerCount ? tr : setRowNumber(tr, i - headerCount + 1)));
    const firstTr = table.inner.indexOf("<w:tr");
    const rebuilt = table.inner.slice(0, firstTr) + out.join("") + "</w:tbl>";
    state.document =
      state.document.slice(0, table.start) + rebuilt + state.document.slice(table.end);
    return {
      operationId: op.id,
      ok: true,
      message:
        `Приложение №${appendix}: добавлено по алфавиту (${inserted.join("; ")}), нумерация обновлена` +
        (skipped.length ? `; пропущено: ${skipped.join("; ")}` : "") +
        (unordered
          ? " — ВНИМАНИЕ: сама таблица упорядочена не по алфавиту, проверьте место вставки"
          : ""),
      orderKey: table.start,
    };
  }

  // ── Изложить предложение или абзац пункта в новой редакции ─────────
  if (op.type === "replace_sentence" || op.type === "replace_paragraph") {
    if (op.payload === undefined) return fail("нет текста новой редакции");
    const point = opPoint(op);
    if (!point) return fail("не указан номер пункта");
    const block = locatePointBlock(
      state.document,
      state.numbering,
      point,
      state.styles,
      searchFrom(op, state.document),
    );
    const withText = block.filter((b) => paragraphText(b.inner).trim().length > 0);
    if (withText.length === 0) return fail(`пункт ${point} не найден в документе`);
    const body = stripLeadingNumber(stripOuterQuotes(op.payload));

    // Абзац пункта: «второй абзац п. 7.6 изложить…». Отрицательный номер —
    // отсчёт с конца, так что −1 это последний абзац.
    if (op.type === "replace_paragraph") {
      const idx = pickIndex(op.paragraphIndex ?? -1, withText.length);
      const span = withText[idx];
      if (!span) return fail(`в п. ${point} нет абзаца № ${op.paragraphIndex}`);
      const keptRefs = footnoteRefRuns(span.inner);
      const old = paragraphText(span.inner).replace(/\s+/g, " ").trim();
      const runs = renderOldPrefix(old, opts) + renderInsertRuns(body, opts) + keptRefs;
      state.document = spliceSpan(state.document, span, replaceParagraphRuns(span.inner, runs));
      return {
        operationId: op.id,
        ok: true,
        message: `п. ${point}: ${ordinalWord(op.paragraphIndex ?? -1, "m")} абзац изложен в новой редакции (абзацев в пункте: ${withText.length})`,
        orderKey: span.start,
        oldText: old,
        newText: body,
      };
    }

    // Предложение. Первое ищем в абзаце-зачине, последнее — в завершающем
    // абзаце пункта: он может идти после перечисления подпунктов.
    const wanted = op.sentenceIndex ?? -1;
    const span = wanted === 1 ? withText[0] : withText[withText.length - 1];
    const plain = paragraphText(span.inner).replace(/\s+/g, " ").trim();
    const parts = sentences(plain);
    if (parts.length === 0) return fail(`пункт ${point} пуст`);
    const sIdx = pickIndex(wanted, parts.length);
    const oldSentence = parts[sIdx];
    if (oldSentence === undefined) return fail(`в п. ${point} нет предложения № ${wanted}`);
    if (plain.includes(body.trim())) {
      return {
        operationId: op.id,
        ok: true,
        message: `п. ${point}: новая редакция предложения уже присутствует`,
        orderKey: span.start,
      };
    }
    const runs = renderOldPrefix(oldSentence, opts) + renderInsertRuns(body, opts);
    const res = replacePhraseRuns(span.inner, oldSentence, runs);
    if (!res.ok)
      return fail(`п. ${point}: ${ordinalWord(wanted)} предложение не удалось выделить — ${res.message}`);
    state.document = spliceSpan(state.document, span, res.xml);
    return {
      operationId: op.id,
      ok: true,
      message:
        `п. ${point}: ${ordinalWord(wanted)} предложение изложено в новой редакции` +
        (block.length > 1 ? ` (пункт из ${block.length} абз.)` : ""),
      orderKey: span.start,
      oldText: oldSentence,
      newText: body,
    };
  }

  // ── Дополнить пункт новым абзацем ──────────────────────────────────
  if (op.type === "append_paragraph") {
    if (op.payload === undefined) return fail("нет текста абзаца");
    const point = opPoint(op);
    if (!point) return fail("не указан номер пункта");
    const block = locatePointBlock(
      state.document,
      state.numbering,
      point,
      state.styles,
      searchFrom(op, state.document),
    );
    const withText = block.filter((b) => paragraphText(b.inner).trim().length > 0);
    const span = withText[withText.length - 1];
    if (!span) return fail(`пункт ${point} не найден в документе`);
    const body = stripLeadingNumber(stripOuterQuotes(op.payload));
    // Абзац внутри пункта своего номера не получает, поэтому автонумерацию с
    // него снимаем — иначе Word посчитает его следующим пунктом.
    const pPr = withoutNumbering(extractPPr(span.inner));
    const newPara = `<w:p>${pPr}${renderInsertRuns(body, opts)}</w:p>`;
    state.document =
      state.document.slice(0, span.end) + newPara + state.document.slice(span.end);
    return {
      operationId: op.id,
      ok: true,
      message: `п. ${point}: добавлен абзац`,
      orderKey: span.end,
    };
  }

  // ── Дополнить пункт предложением ───────────────────────────────────
  if (op.type === "append_sentence") {
    if (op.payload === undefined) return fail("нет текста предложения");
    const point = opPoint(op);
    if (!point) return fail("не указан номер пункта");
    const block = locatePointBlock(
      state.document,
      state.numbering,
      point,
      state.styles,
      searchFrom(op, state.document),
    );
    const withText = block.filter((b) => paragraphText(b.inner).trim().length > 0);
    const span = withText[withText.length - 1];
    if (!span) return fail(`пункт ${point} не найден в документе`);
    const body = stripLeadingNumber(stripOuterQuotes(op.payload));
    const plain = paragraphText(span.inner).replace(/\s+/g, " ");
    if (plain.includes(body.trim())) {
      return {
        operationId: op.id,
        ok: true,
        message: `п. ${point}: предложение уже присутствует`,
        orderKey: span.start,
      };
    }
    // «В конце второго предложения дополнить словами» — текст идёт в конец
    // именно этого предложения, а не всего пункта.
    if (op.sentenceIndex !== undefined) {
      const parts = sentences(plain.trim());
      const idx = pickIndex(op.sentenceIndex, parts.length);
      const target = parts[idx];
      if (!target) return fail(`в п. ${point} нет предложения № ${op.sentenceIndex}`);
      const res = replacePhraseRuns(
        span.inner,
        target,
        renderInsertRuns(target, opts).replace(/<w:color w:val="[^"]*"\/>/, "") +
          renderInsertRuns(body, opts),
      );
      if (!res.ok) return fail(`п. ${point}: ${res.message}`);
      state.document = spliceSpan(state.document, span, res.xml);
      return {
        operationId: op.id,
        ok: true,
        message: `п. ${point}: ${ordinalWord(op.sentenceIndex)} предложение дополнено`,
        orderKey: span.start,
      };
    }
    const closing = span.inner.lastIndexOf("</w:p>");
    const rebuilt =
      span.inner.slice(0, closing) + renderInsertRuns(" " + body, opts) + span.inner.slice(closing);
    state.document = spliceSpan(state.document, span, rebuilt);
    return {
      operationId: op.id,
      ok: true,
      message: `п. ${point}: дополнен предложением`,
      orderKey: span.start,
    };
  }

  // ── Заменить / удалить слова внутри пункта ─────────────────────────
  if (op.type === "replace_words" || op.type === "delete_words") {
    if (!op.find) return fail("не указано, какие слова менять");
    const point = opPoint(op);
    const span = point ? locatePointSpan(state.document, state.numbering, point, state.styles, searchFrom(op, state.document)) : null;
    if (point && !span) return fail(`пункт ${point} не найден в документе`);
    const scope = span ? span.inner : state.document;
    const occurrence = op.anchor ? phraseOccurrenceAfter(scope, op.anchor, op.find) : 0;
    if (occurrence === null)
      return fail(`в п. ${point}: слова «${op.find}» после «${op.anchor}» не найдены`);
    const hits = countPhrase(scope, op.find);
    if (hits === 0) {
      // Возможно, правка уже внесена в этой редакции.
      const replacement = op.payload ? stripOuterQuotes(op.payload) : "";
      if (replacement && countPhrase(scope, replacement) > 0) {
        return {
          operationId: op.id,
          ok: true,
          message: `п. ${point}: «${replacement}» уже стоит вместо «${op.find}» — правка не требуется`,
          orderKey: span ? span.start : 0,
        };
      }
      return fail(`слова «${op.find}» не найдены${point ? ` в п. ${point}` : ""}`);
    }
    if (!span && hits > 1)
      return fail(`слова «${op.find}» встречаются ${hits} раз, а пункт не указан — правка неоднозначна`);

    const del = renderOldInline(op.find, opts);
    const runs =
      op.type === "delete_words"
        ? del
        : del + renderInsertRuns(stripOuterQuotes(op.payload ?? ""), opts);
    const res = replacePhraseRuns(scope, op.find, runs, occurrence);
    if (!res.ok) return fail(res.message);
    state.document = span ? spliceSpan(state.document, span, res.xml) : res.xml;
    const many = hits > 1 && span ? ` (в пункте ${hits} вхождений, изменено одно)` : "";
    return {
      operationId: op.id,
      ok: true,
      message:
        (op.type === "delete_words"
          ? `удалены слова «${op.find}»`
          : `«${op.find}» заменено на «${stripOuterQuotes(op.payload ?? "")}»`) +
        (point ? ` в п. ${point}` : "") +
        many,
      orderKey: span ? span.start : res.orderKey,
      oldText: op.find,
      newText: op.type === "delete_words" ? undefined : stripOuterQuotes(op.payload ?? ""),
    };
  }

  // ── Исключить пункт (последующие перенумеровываются) ───────────────
  if (op.type === "delete_point") {
    const point = opPoint(op);
    if (!point) return fail("не указан номер исключаемого пункта");
    const span = locatePointSpan(state.document, state.numbering, point, state.styles, searchFrom(op, state.document));
    if (!span) return fail(`пункт ${point} не найден в документе`);
    const plain = paragraphText(span.inner).replace(/\s+/g, " ").trim();
    // Снимаем автонумерацию: тогда Word перенумерует последующие пункты, а сам
    // текст остаётся зачёркнутым — читатель видит, что именно исключено.
    const pPr = withoutNumbering(extractPPr(span.inner));
    const openMatch = span.inner.match(/^<w:p(?:\s[^>]*)?>/);
    const open = openMatch ? openMatch[0] : "<w:p>";
    const body = stripLeadingNumber(plain);
    const rebuilt = `${open}${pPr}${renderDeleteRuns(`${point}. ${body}`, opts)}</w:p>`;
    state.document = spliceSpan(state.document, span, rebuilt);
    return {
      operationId: op.id,
      ok: true,
      message: `пункт ${point} исключён (зачёркнут, последующие перенумеровываются автоматически)`,
      orderKey: span.start,
      oldText: plain,
    };
  }

  // ── Замена слов по всему тексту ────────────────────────────────────
  if (op.type === "replace_words_global") {
    if (!op.find) return fail("не указано, что заменять");
    const replacement = stripOuterQuotes(op.payload ?? "");
    const total = countPhrase(state.document, op.find);
    if (total === 0) {
      const already = countPhrase(state.document, replacement);
      if (replacement && already > 0) {
        return {
          operationId: op.id,
          ok: true,
          message: `«${replacement}» уже стоит вместо «${op.find}» (${already} мест) — правка не требуется`,
          orderKey: 0,
        };
      }
      return fail(`«${op.find}» в тексте не найдено`);
    }
    const runs = renderOldInline(op.find, opts) + renderInsertRuns(replacement, opts);
    let done = 0;
    let skipped = 0;
    // Идём с конца: каждая замена меняет длину строки, и позиции более ранних
    // вхождений от этого не сдвигаются.
    for (let i = total - 1; i >= 0; i--) {
      const res = replacePhraseRuns(state.document, op.find, runs, i);
      if (res.ok) {
        state.document = res.xml;
        done++;
      } else {
        skipped++;
      }
    }
    if (done === 0) return fail(`замены не удались: ${skipped} мест пересекают сноски или объекты`);
    return {
      operationId: op.id,
      ok: true,
      message:
        `по всему тексту заменено «${op.find}» → «${replacement}»: ${done} мест` +
        (skipped ? `; пропущено ${skipped} (пересекают сноску или объект — проверьте вручную)` : ""),
      orderKey: 0,
      oldText: op.find,
      newText: replacement,
    };
  }

  // ── Исключить абзац пункта ─────────────────────────────────────────
  if (op.type === "delete_paragraph") {
    const point = opPoint(op);
    if (!point) return fail("не указан номер пункта");
    const block = locatePointBlock(
      state.document,
      state.numbering,
      point,
      state.styles,
      searchFrom(op, state.document),
    );
    const withText = block.filter((b) => paragraphText(b.inner).trim().length > 0);
    if (!withText.length) return fail(`пункт ${point} не найден в документе`);
    const idx = pickIndex(op.paragraphIndex ?? -1, withText.length);
    const span = withText[idx];
    if (!span) return fail(`в п. ${point} нет абзаца № ${op.paragraphIndex}`);
    const plain = paragraphText(span.inner).replace(/\s+/g, " ").trim();
    const pPr = withoutNumbering(extractPPr(span.inner));
    const openMatch = span.inner.match(/^<w:p(?:\s[^>]*)?>/);
    const rebuilt = `${openMatch ? openMatch[0] : "<w:p>"}${pPr}${renderDeleteRuns(plain, opts)}</w:p>`;
    state.document = spliceSpan(state.document, span, rebuilt);
    return {
      operationId: op.id,
      ok: true,
      message: `п. ${point}: ${ordinalWord(op.paragraphIndex ?? -1, "m")} абзац исключён (зачёркнут; всего абзацев было ${withText.length})`,
      orderKey: span.start,
      oldText: plain,
    };
  }

  // ── Исключить сноску ───────────────────────────────────────────────
  if (op.type === "delete_footnote") {
    const number = resolveFootnoteNumber(op, state);
    if (number === null) return fail("не удалось определить, какую сноску исключить");
    const idx = indexFootnotes(state.document);
    const pos = idx.displayToBodyPos.get(number);
    if (pos === undefined) return fail(`сноска № ${number} не найдена в тексте`);
    // Убираем ссылку из текста: без неё Word перестаёт показывать сноску и
    // перенумеровывает последующие. Сам текст сноски остаётся в footnotes.xml —
    // он не отображается, но и не теряется безвозвратно.
    const runStart = state.document.lastIndexOf("<w:r", pos);
    const runEnd = state.document.indexOf("</w:r>", pos);
    if (runStart < 0 || runEnd < 0) return fail(`не удалось выделить ссылку на сноску № ${number}`);
    const oldFnText = state.footnotes
      ? paragraphText(findFootnoteById(state.footnotes, idx.displayToId.get(number)!)?.inner ?? "").replace(/\s+/g, " ").trim()
      : undefined;
    state.document =
      state.document.slice(0, runStart) + state.document.slice(runEnd + "</w:r>".length);
    return {
      operationId: op.id,
      ok: true,
      message: `сноска № ${number} исключена (последующие перенумеровываются автоматически)`,
      orderKey: runStart,
      oldText: oldFnText,
    };
  }

  // ── Требует ручной обработки ───────────────────────────────────────
  if (op.type === "manual") {
    const marked = insertManualMarker(op, state);
    return {
      operationId: op.id,
      ok: false,
      message:
        `требует ручной обработки: ${op.note ?? op.rawText.slice(0, 80)}` +
        (marked ? " (место отмечено в документе)" : ""),
      orderKey: marked ?? Number.MAX_SAFE_INTEGER,
    };
  }

  return fail(`тип операции не поддержан: ${op.type}`);
}

/**
 * Заметная метка «здесь нужна ручная правка» прямо в тексте Оферты.
 *
 * Оператор работает с документом, а не со списком на экране, и правку, которую
 * движок не может внести сам, легко пропустить. Настоящий комментарий Word тут
 * не годится: часть редакций Оферты приходит вообще без word/comments.xml, и
 * добавление этой части — лишний риск испортить файл. Жёлтая заливка с красным
 * текстом видна сразу и работает одинаково во всех редакциях.
 */
function manualMarkerText(op: Operation): string[] {
  const what = op.note ?? op.rawText.slice(0, 300);
  const rows = (op.rows ?? [])
    .map((r) => r.map((c) => c.replace(/\s+/g, " ").trim()).filter(Boolean).join(" | "))
    .filter(Boolean);
  return [`ТРЕБУЕТСЯ РУЧНАЯ ПРАВКА: ${what}`, ...rows.map((r) => `→ ${r}`)];
}

/**
 * Абзац-метка, к которому привязывается комментарий.
 *
 * Комментарий Word виден только при открытой области рецензирования, поэтому
 * цветная метка остаётся: она заметна в любом случае и служит якорем, к
 * которому комментарий привязан. `commentId` — номер комментария, если он был
 * заведён.
 */
function manualMarkerParagraph(op: Operation, commentId: number | null): string {
  const runs = manualMarkerText(op)
    .map(
      (line, i) =>
        (i > 0 ? `<w:r><w:rPr><w:b/><w:color w:val="C00000"/><w:highlight w:val="yellow"/></w:rPr><w:br/></w:r>` : "") +
        `<w:r><w:rPr><w:b/><w:color w:val="C00000"/><w:highlight w:val="yellow"/></w:rPr>` +
        `<w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r>`,
    )
    .join("");
  const body =
    commentId === null
      ? runs
      : commentRangeStart(commentId) + runs + commentRangeEnd(commentId);
  return `<w:p><w:pPr><w:spacing w:before="120" w:after="120"/></w:pPr>${body}</w:p>`;
}

/**
 * Поставить метку у цели ручной правки. Возвращает позицию вставки либо null,
 * если цель в Оферте не найдена (например, правка относится к другому
 * документу Альбома форм — тогда метке в Оферте не место).
 */
function insertManualMarker(op: Operation, state: ApplyState): number | null {
  let at: number | null = null;
  if (op.target.kind === "appendix_table") {
    const table = findAppendixTable(state.document, op.target.appendix);
    if (table) at = table.start;
  } else if (op.target.kind === "preamble") {
    const span = locatePreamble(state.document, state.numbering, state.styles);
    if (span) at = span.end;
  } else {
    const point = opPoint(op);
    // «—» ставится, когда пункт неизвестен: цели в Оферте нет.
    if (point && point !== "—") {
      const span = locatePointSpan(state.document, state.numbering, point, state.styles);
      if (span) at = span.end;
    }
  }
  if (at === null) return null;
  const commentId = state.comments
    ? addComment(state.comments, manualMarkerText(op).join("\n"))
    : null;
  const marker = manualMarkerParagraph(op, commentId);
  state.document = state.document.slice(0, at) + marker + state.document.slice(at);
  return at;
}

/**
 * Порядок применения операций. Вынесен отдельно, потому что предпросмотр
 * обязан прогонять операции в ТОМ ЖЕ порядке: иначе оператор увидит одно, а в
 * файл попадёт другое — ровно та рассинхронизация, из-за которой правка пункта
 * оказывалась в документе, но не показывалась на шаге «Проверка».
 */
export function applicationOrder(operations: Operation[]): number[] {
  const isNormalizing = (op: Operation) => op.type === "sort_table_alpha";
  const order: number[] = [];
  for (let i = 0; i < operations.length; i++) {
    if (!isNormalizing(operations[i])) order.push(i);
  }
  for (let i = 0; i < operations.length; i++) {
    if (isNormalizing(operations[i])) order.push(i);
  }
  return order;
}

/** Применить набор операций к Оферте, вернуть байты docx и отчёт. */
export async function applyOperations(
  offer: DocxParts,
  operations: Operation[],
  opts: BuildOptions,
): Promise<{ offerDocx: Uint8Array; results: ApplyResult[] }> {
  resetInsCounter();
  const state: ApplyState = {
    document: offer.document,
    footnotes: offer.footnotes,
    numbering: offer.numbering,
    styles: offer.styles,
    comments: prepareComments(offer.comments, offer.rels, offer.contentTypes),
  };
  const results: ApplyResult[] = [];
  // Порядок применения:
  //  1) В ПОРЯДКЕ ДОКУМЕНТА «Изменения», сверху вниз. Это не косметика:
  //     инструкции внутри пакета ссылаются на нумерацию, которая получается
  //     ПОСЛЕ предыдущих инструкций. Документ добавляет п. 7.3 «с последующей
  //     перенумерацией», и следующая его же строка про «п. 7.6» имеет в виду
  //     пункт, который до этой вставки был 7.5. Каждая операция ищет цель
  //     заново в текущем состоянии документа, поэтому сдвиг позиций безопасен;
  //  2) нормализующие операции (алфавитная пересортировка) — В САМОМ КОНЦЕ:
  //     они приводят таблицу в порядок уже ПОСЛЕ всех замен и добавлений
  //     строк, иначе переименованные строки нарушат алфавитный порядок.
  const order = applicationOrder(operations);

  const slots: (ApplyResult | undefined)[] = new Array(operations.length);
  for (const i of order) {
    slots[i] = applyOneOp(operations[i], state, opts);
  }
  for (const r of slots) if (r) results.push(r);
  // Часть комментариев (и связь с типом) пишем только если комментарии реально
  // появились: пустая часть в пакете Word не нужна.
  const c = state.comments;
  const offerDocx = await saveDocx(offer, {
    document: state.document,
    footnotes: state.footnotes ?? undefined,
    comments: c?.touched ? c.xml : undefined,
    rels: c?.touched ? c.rels : undefined,
    contentTypes: c?.touched ? c.contentTypes : undefined,
  });
  return { offerDocx, results };
}
