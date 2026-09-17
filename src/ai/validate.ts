// Проверка ответа модели.
//
// Ответ модели — это НЕДОВЕРЕННЫЕ данные, даже когда он выглядит безупречно.
// Между моделью и документом стоят две преграды:
//   1) схема — операция обязана быть синтаксически корректной и полной
//      (у insert_after есть anchor, у replace_words есть find, и т.д.);
//   2) сухой прогон — операция обязана НАЙТИ свою цель в реальной Оферте
//      (см. dryRun в hybrid.ts). Пункт 12.7, которого нет, схему пройдёт,
//      а прогон — нет.
// Всё, что не прошло, не выбрасывается, а превращается в `manual`: молча
// терять правку нельзя, оператор должен увидеть каждую строку.
import type { OpTarget, OpType, Operation } from "@/engine/types";

const OP_TYPES = new Set<OpType>([
  "insert_after",
  "replace",
  "replace_footnote",
  "add_footnote",
  "insert_point",
  "append_table_rows",
  "replace_table_rows",
  "sort_table_alpha",
  "insert_table_row_alpha",
  "insert_before",
  "replace_sentence",
  "replace_paragraph",
  "append_sentence",
  "append_paragraph",
  "replace_words",
  "delete_words",
  "replace_words_global",
  "delete_paragraph",
  "delete_footnote",
  "delete_point",
  "manual",
]);

/** Операции, без payload бессмысленные: вставлять/заменять нечем. */
const NEEDS_PAYLOAD = new Set<OpType>([
  "insert_after",
  "insert_before",
  "replace",
  "replace_footnote",
  "add_footnote",
  "insert_point",
  "replace_sentence",
  "replace_paragraph",
  "append_sentence",
  "append_paragraph",
  "replace_words",
  "replace_words_global",
]);

/** Операции, которым нужен якорь в тексте Оферты. */
const NEEDS_ANCHOR = new Set<OpType>(["insert_after", "insert_before"]);

/** Операции, которым нужен искомый фрагмент. */
const NEEDS_FIND = new Set<OpType>([
  "replace_words",
  "delete_words",
  "replace_words_global",
]);

/** Операции, которым нужны строки таблицы. */
const NEEDS_ROWS = new Set<OpType>([
  "append_table_rows",
  "replace_table_rows",
  "insert_table_row_alpha",
]);

export interface ValidationIssue {
  field: string;
  message: string;
}

export interface ValidatedOperation {
  ok: boolean;
  /** Операция без служебных полей (id/sourceDoc/rawText проставит вызывающий). */
  draft?: Omit<Operation, "id" | "sourceDoc" | "rawText">;
  issues: ValidationIssue[];
}

function str(v: unknown): string | undefined {
  return typeof v === "string" && v.trim() ? v : undefined;
}

function int(v: unknown): number | undefined {
  if (typeof v === "number" && Number.isFinite(v)) return Math.trunc(v);
  if (typeof v === "string" && /^-?\d+$/.test(v.trim())) return parseInt(v, 10);
  return undefined;
}

/**
 * Снять внешние кавычки-ёлочки, если модель их всё-таки оставила.
 *
 * Промпт просит этого не делать, но полагаться на дисциплину модели нельзя:
 * лишняя пара кавычек в payload — это лишняя пара кавычек в тексте договора.
 * Снимаем только ВНЕШНЮЮ пару и только когда она обрамляет строку целиком,
 * иначе пострадает «ООО «Ромашка»» — там кавычки часть названия.
 */
export function stripOuterQuotes(s: string): string {
  const t = s.trim();
  if (!t.startsWith("«") || !t.endsWith("»")) return t;
  let depth = 0;
  for (let i = 0; i < t.length; i++) {
    if (t[i] === "«") depth++;
    else if (t[i] === "»") {
      depth--;
      // Внешняя кавычка закрылась раньше конца строки — значит, это не обёртка
      // всей строки, а начало вложенной пары: «А», «Б» трогать нельзя.
      if (depth === 0 && i < t.length - 1) return t;
    }
  }
  return depth === 0 ? t.slice(1, -1) : t;
}

function parseTarget(v: unknown, issues: ValidationIssue[]): OpTarget | null {
  if (!v || typeof v !== "object") {
    issues.push({ field: "target", message: "цель правки не указана" });
    return null;
  }
  const t = v as Record<string, unknown>;
  const kind = str(t.kind);
  switch (kind) {
    case "footnote": {
      const number = int(t.number);
      const atPoint = str(t.atPoint);
      if (number == null && !atPoint) {
        issues.push({ field: "target.number", message: "сноска не задана ни номером, ни пунктом" });
        return null;
      }
      // Внутренний контракт движка требует number; для «сноски к пункту 5.3»
      // номер подставит локатор, здесь достаточно нуля-заглушки.
      return { kind: "footnote", number: number ?? 0, label: str(t.label), atPoint };
    }
    case "term": {
      const term = str(t.term);
      if (!term) {
        issues.push({ field: "target.term", message: "термин не указан" });
        return null;
      }
      return { kind: "term", term, point: str(t.point), section: str(t.section) };
    }
    case "point": {
      const point = str(t.point);
      if (!point) {
        issues.push({ field: "target.point", message: "номер пункта не указан" });
        return null;
      }
      return { kind: "point", point, section: str(t.section), heading: str(t.heading) };
    }
    case "preamble":
      return { kind: "preamble" };
    case "appendix_point": {
      const appendix = str(t.appendix);
      const point = str(t.point);
      if (!appendix || !point) {
        issues.push({
          field: "target.appendix",
          message: "для пункта приложения нужны и номер приложения, и номер пункта",
        });
        return null;
      }
      return { kind: "appendix_point", appendix, point };
    }
    case "appendix_table": {
      const appendix = str(t.appendix);
      if (!appendix) {
        issues.push({ field: "target.appendix", message: "не указан номер приложения" });
        return null;
      }
      return { kind: "appendix_table", appendix, point: str(t.point) };
    }
    default:
      issues.push({ field: "target.kind", message: `неизвестный вид цели: ${String(t.kind)}` });
      return null;
  }
}

function parseRows(v: unknown): string[][] | undefined {
  if (!Array.isArray(v)) return undefined;
  const rows = v
    .filter(Array.isArray)
    .map((r) => (r as unknown[]).map((c) => (c == null ? "" : String(c).trim())));
  return rows.length ? rows : undefined;
}

/** Проверить одну операцию из ответа модели. */
export function validateOperation(raw: unknown): ValidatedOperation {
  const issues: ValidationIssue[] = [];
  if (!raw || typeof raw !== "object") {
    return { ok: false, issues: [{ field: "", message: "операция не является объектом" }] };
  }
  const o = raw as Record<string, unknown>;

  const type = str(o.type) as OpType | undefined;
  if (!type || !OP_TYPES.has(type)) {
    return {
      ok: false,
      issues: [{ field: "type", message: `неизвестный тип операции: ${String(o.type)}` }],
    };
  }

  const target = parseTarget(o.target, issues);
  if (!target) return { ok: false, issues };

  const payload = str(o.payload) ? stripOuterQuotes(String(o.payload)) : undefined;
  const anchor = str(o.anchor) ? stripOuterQuotes(String(o.anchor)) : undefined;
  const find = str(o.find) ? stripOuterQuotes(String(o.find)) : undefined;
  const rows = parseRows(o.rows);

  if (type !== "manual") {
    if (NEEDS_PAYLOAD.has(type) && !payload) {
      issues.push({ field: "payload", message: `для «${type}» не найден текст правки` });
    }
    if (NEEDS_ANCHOR.has(type) && !anchor) {
      issues.push({ field: "anchor", message: `для «${type}» не указаны слова-якорь` });
    }
    if (NEEDS_FIND.has(type) && !find) {
      issues.push({ field: "find", message: `для «${type}» не указан искомый фрагмент` });
    }
    if (NEEDS_ROWS.has(type) && !rows) {
      issues.push({ field: "rows", message: `для «${type}» не найдены строки таблицы` });
    }
    // Замена «А на Б» без исходного «А» заменит не то и не там.
    if (type === "replace_words" && find && payload && find === payload) {
      issues.push({ field: "find", message: "искомый и заменяющий тексты совпадают" });
    }
  }

  if (issues.length) return { ok: false, issues };

  const confidence = typeof o.confidence === "number" ? clamp01(o.confidence) : 0.7;
  return {
    ok: true,
    issues: [],
    draft: {
      type,
      target,
      payload,
      anchor,
      find,
      rows,
      rowNumbers: Array.isArray(o.rowNumbers)
        ? (o.rowNumbers.map(int).filter((n): n is number => n != null) ?? undefined)
        : undefined,
      sentenceIndex: int(o.sentenceIndex),
      paragraphIndex: int(o.paragraphIndex),
      nameColumn: int(o.nameColumn),
      note: str(o.note),
      confidence,
      warnings: undefined,
    },
  };
}

function clamp01(n: number): number {
  return Math.min(1, Math.max(0, n));
}

/** Разобрать ответ модели целиком. */
export function parseAiResponse(data: unknown): {
  operations: unknown[];
  reasoning?: string;
  confidence?: number;
} {
  if (!data || typeof data !== "object") return { operations: [] };
  const d = data as Record<string, unknown>;
  const ops = Array.isArray(d.operations)
    ? d.operations
    : // Модель иногда отдаёт одну операцию без обёртки в массив.
      d.type
      ? [d]
      : [];
  return {
    operations: ops,
    reasoning: str(d.reasoning),
    confidence: typeof d.confidence === "number" ? clamp01(d.confidence) : undefined,
  };
}
