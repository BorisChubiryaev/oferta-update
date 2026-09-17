// Гибридный разбор: алгоритм ведёт, ИИ подстраховывает, движок проверяет.
//
// Порядок именно такой, а не «спросить модель обо всём»:
//
//   1. Детерминированный разбор обрабатывает документ целиком. Формулировки в
//      нормативных документах устойчивы, и правила разбирают большую их часть
//      бесплатно, мгновенно и ВОСПРОИЗВОДИМО — один и тот же документ всегда
//      даёт один и тот же результат. Терять это ради единообразия «всё через
//      ИИ» незачем.
//   2. Модель зовут точечно: туда, где алгоритм сдался (`manual`), где он не
//      уверен (низкая confidence), и — если оператор попросил — вторым мнением
//      на спорные места.
//   3. ЛЮБАЯ операция, пришедшая от модели, проходит сухой прогон по реальной
//      Оферте. Ссылка на несуществующий пункт, якорь, которого нет в тексте,
//      сноска за пределами нумерации — всё это отсеивается здесь, а не в
//      готовом документе.
//
// Что бы ни ответила модель, текст Оферты меняет движок — по координатам,
// прошедшим проверку.
import { loadDocx } from "@/engine/docx";
import { applyOneOp } from "@/engine/apply";
import { analyzeChangeDoc } from "@/engine/offline";
import { documentBlocks, tables as docTables } from "@/engine/text";
import { buildOfferDigest, digestToPrompt } from "@/engine/offer-digest";
import { validateOperation } from "./validate";
import type { BuildOptions, Operation, OpTarget } from "@/engine/types";
import type {
  AiParseRequest,
  AiParseResponse,
  HybridOperation,
  HybridResult,
  Provenance,
} from "./types";

/** Ниже этого порога разбор считается неуверенным и отдаётся на проверку ИИ. */
const LOW_CONFIDENCE = 0.6;

export interface HybridInput {
  offer: Uint8Array;
  changeDocs: { name: string; data: Uint8Array }[];
}

export interface HybridSettings {
  /** Звать ли модель вообще. Выключено — работает прежний алгоритмический режим. */
  useAi: boolean;
  /**
   * Проверять моделью и те правки, с которыми алгоритм справился уверенно.
   * Дороже и медленнее, зато ловит случай «правило сработало, но не то».
   */
  verifyConfident: boolean;
  /** Сообщать о ходе работы (для индикатора в интерфейсе). */
  onProgress?: (done: number, total: number, stage: string) => void;
}

/** Обращение к модели вынесено в параметр — так конвейер тестируется без сети. */
export type AiCaller = (req: AiParseRequest) => Promise<AiParseResponse>;

/** Боевой вызов: через серверный маршрут, ключ остаётся на сервере. */
export const httpAiCaller: AiCaller = async (req) => {
  const res = await fetch("/api/ai/parse", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(req),
  });
  if (!res.ok) {
    const body = (await res.json().catch(() => ({}))) as { error?: string };
    throw new Error(body.error ?? `Сервис ИИ вернул ${res.status}`);
  }
  return (await res.json()) as AiParseResponse;
};

/**
 * Сухой прогон операции по копии Оферты.
 *
 * Это главная проверка предложений модели, и она не «почти бесплатная»
 * эвристика, а тот же самый код, который потом реально правит документ:
 * applyOneOp ищет цель, якорь и текст ровно так же, как при сборке. Если он
 * говорит «пункт 12.7 не найден» — значит, при сборке не найдёт тоже.
 */
function dryRun(
  op: Operation,
  offerParts: Awaited<ReturnType<typeof loadDocx>>,
  opts: BuildOptions,
): { ok: boolean; message: string; orderKey: number } {
  // Состояние — копия строк: applyOneOp возвращает новые строки, не мутируя
  // исходные, поэтому прогон не портит документ для следующих проверок.
  const state = {
    document: offerParts.document,
    footnotes: offerParts.footnotes,
    numbering: offerParts.numbering,
    styles: offerParts.styles,
  };
  try {
    const r = applyOneOp(op, state, opts);
    return { ok: r.ok, message: r.message, orderKey: r.orderKey };
  } catch (e) {
    return {
      ok: false,
      message: e instanceof Error ? e.message : "ошибка при проверке операции",
      orderKey: Number.MAX_SAFE_INTEGER,
    };
  }
}

let seq = 0;
function nextId(src: string): string {
  seq += 1;
  return `${src}#ai${seq}`;
}

/** Черновик от модели -> полноценная операция. */
function toOperation(
  draft: Omit<Operation, "id" | "sourceDoc" | "rawText">,
  sourceDoc: string,
  rawText: string,
  warnings: string[],
): Operation {
  return {
    ...draft,
    id: nextId(sourceDoc),
    sourceDoc,
    rawText,
    warnings: warnings.length ? warnings : draft.warnings,
  };
}

/** Операция «нужен человек» — на случай, когда не помогли ни правила, ни модель. */
function manualOp(sourceDoc: string, rawText: string, note: string): Operation {
  return {
    id: nextId(sourceDoc),
    sourceDoc,
    type: "manual",
    target: { kind: "point", point: "—" },
    note,
    confidence: 0.3,
    rawText,
    warnings: [note],
  };
}

/**
 * Совпадают ли две цели.
 *
 * Сравнивать через JSON.stringify нельзя: порядок ключей у объекта, собранного
 * правилом разбора, и у объекта, пересобранного проверкой ответа модели, разный,
 * и одинаковые по смыслу цели выглядели бы как расхождение. Оператор в итоге
 * получал бы «расхождение» там, где алгоритм и модель согласны, — то есть ровно
 * там, где проверять как раз не нужно.
 */
function sameTarget(a: OpTarget, b: OpTarget): boolean {
  if (a.kind !== b.kind) return false;
  const norm = (s?: string) => (s ?? "").trim().toLowerCase().replace(/[.\s]+$/, "");
  switch (a.kind) {
    case "preamble":
      return true;
    case "footnote": {
      const o = b as Extract<OpTarget, { kind: "footnote" }>;
      // Сноска, заданная пунктом, номера ещё не знает (его подставит локатор),
      // поэтому по номеру сравниваем только когда он задан у обеих сторон.
      if (a.atPoint || o.atPoint) return norm(a.atPoint) === norm(o.atPoint);
      return a.number === o.number;
    }
    case "term": {
      const o = b as Extract<OpTarget, { kind: "term" }>;
      return norm(a.term) === norm(o.term) && norm(a.point) === norm(o.point);
    }
    case "point": {
      const o = b as Extract<OpTarget, { kind: "point" }>;
      // Заголовок раздела — подсказка оператору, а не координата: расходиться
      // в нём алгоритм и модель могут сколько угодно.
      return norm(a.point) === norm(o.point);
    }
    case "appendix_point": {
      const o = b as Extract<OpTarget, { kind: "appendix_point" }>;
      return norm(a.appendix) === norm(o.appendix) && norm(a.point) === norm(o.point);
    }
    case "appendix_table": {
      const o = b as Extract<OpTarget, { kind: "appendix_table" }>;
      return norm(a.appendix) === norm(o.appendix) && norm(a.point) === norm(o.point);
    }
  }
}

/** Совпадают ли два прочтения по сути (тип + цель + текст правки). */
function sameReading(a: Operation, b: Operation): boolean {
  if (a.type !== b.type) return false;
  if (!sameTarget(a.target, b.target)) return false;
  const norm = (s?: string) => (s ?? "").replace(/\s+/g, " ").trim().toLowerCase();
  return norm(a.payload) === norm(b.payload) && norm(a.anchor) === norm(b.anchor);
}

/** Нужна ли этой правке помощь модели. */
function needsAi(ops: Operation[], verifyConfident: boolean): boolean {
  if (!ops.length) return true;
  if (verifyConfident) return true;
  return ops.some((o) => o.type === "manual" || o.confidence < LOW_CONFIDENCE);
}

export async function runHybrid(
  input: HybridInput,
  settings: HybridSettings,
  callAi: AiCaller = httpAiCaller,
): Promise<HybridResult> {
  seq = 0;
  const opts: BuildOptions = {};
  const offerParts = await loadDocx(input.offer);
  const digest = digestToPrompt(buildOfferDigest(offerParts));

  // Шаг 1 — детерминированный разбор всех документов «Изменения».
  settings.onProgress?.(0, 1, "Разбор документов алгоритмом");
  const units: {
    sourceDoc: string;
    index: number;
    text: string;
    tables: string[][][];
    operations: Operation[];
  }[] = [];
  for (const cd of input.changeDocs) {
    const parts = await loadDocx(cd.data);
    const analysis = analyzeChangeDoc(
      documentBlocks(parts.document),
      docTables(parts.document),
      cd.name,
    );
    for (const u of analysis.units) {
      units.push({ sourceDoc: cd.name, ...u });
    }
  }

  const out: HybridOperation[] = [];
  const aiErrors: string[] = [];
  let aiUsed = false;

  // Шаг 2 — точечная помощь модели там, где она нужна.
  const queue = new Set(
    settings.useAi ? units.filter((u) => needsAi(u.operations, settings.verifyConfident)) : [],
  );
  let done = 0;

  for (const unit of units) {
    const wantsAi = queue.has(unit);
    if (!wantsAi) {
      // Алгоритм справился и второе мнение не запрашивали.
      for (const op of unit.operations) {
        const run = dryRun(op, offerParts, opts);
        out.push({
          op,
          provenance: op.type === "manual" ? "manual" : "algorithm",
          unitIndex: unit.index,
          dryRun: { ok: run.ok, message: run.message },
        });
      }
      continue;
    }

    settings.onProgress?.(done, queue.size, `Запрос к ИИ: правка ${done + 1} из ${queue.size}`);
    done += 1;

    const algorithmOps = unit.operations.filter((o) => o.type !== "manual");
    const mode = algorithmOps.length ? "verify" : "parse";
    let response: AiParseResponse | null = null;
    try {
      aiUsed = true;
      response = await callAi({
        instruction: unit.text,
        tables: unit.tables.length ? unit.tables.slice(0, 3) : undefined,
        offerDigest: digest,
        algorithmDraft: mode === "verify" ? algorithmOps.map(stripInternals) : undefined,
        mode,
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      // Сбой ИИ не должен ронять всю сборку: возвращаемся к тому, что дал
      // алгоритм, и честно сообщаем оператору, что второго мнения не было.
      if (!aiErrors.includes(msg)) aiErrors.push(msg);
      for (const op of unit.operations) {
        const run = dryRun(op, offerParts, opts);
        out.push({
          op: {
            ...op,
            warnings: [...(op.warnings ?? []), "ИИ недоступен — результат только алгоритмический"],
          },
          provenance: op.type === "manual" ? "manual" : "algorithm",
          unitIndex: unit.index,
          dryRun: { ok: run.ok, message: run.message },
        });
      }
      continue;
    }

    // Шаг 3 — проверка предложений модели: схема, затем сухой прогон.
    const accepted: Operation[] = [];
    const rejections: string[] = [];
    for (const rawOp of response.operations) {
      const v = validateOperation(rawOp);
      if (!v.ok || !v.draft) {
        rejections.push(v.issues.map((i) => `${i.field}: ${i.message}`).join("; "));
        continue;
      }
      if (v.draft.type === "manual") {
        rejections.push(v.draft.note ?? "модель не смогла разобрать инструкцию однозначно");
        continue;
      }
      const candidate = toOperation(v.draft, unit.sourceDoc, unit.text, []);
      const run = dryRun(candidate, offerParts, opts);
      if (!run.ok) {
        // Схему прошло, а цели в документе нет — самый частый вид ошибки модели.
        rejections.push(`предложение ИИ не применяется к Оферте: ${run.message}`);
        continue;
      }
      accepted.push(candidate);
    }

    if (!accepted.length) {
      // Модель не помогла. Если у алгоритма что-то было — оставляем его.
      if (algorithmOps.length) {
        for (const op of algorithmOps) {
          const run = dryRun(op, offerParts, opts);
          out.push({
            op,
            provenance: "algorithm",
            unitIndex: unit.index,
            aiReasoning: response.reasoning,
            dryRun: { ok: run.ok, message: run.message },
          });
        }
      } else {
        const note = rejections.length
          ? `ни алгоритм, ни ИИ не дали применимой правки (${rejections[0]})`
          : "ни алгоритм, ни ИИ не распознали формулировку";
        out.push({
          op: manualOp(unit.sourceDoc, unit.text, note),
          provenance: "manual",
          unitIndex: unit.index,
          aiReasoning: response.reasoning,
        });
      }
      continue;
    }

    // Модель дала применимые операции. Сверяем с прочтением алгоритма.
    for (let i = 0; i < accepted.length; i++) {
      const aiOp = accepted[i];
      const twin = algorithmOps[i];
      const run = dryRun(aiOp, offerParts, opts);
      if (!twin) {
        out.push({
          op: aiOp,
          provenance: "ai",
          unitIndex: unit.index,
          aiReasoning: response.reasoning,
          dryRun: { ok: run.ok, message: run.message },
        });
        continue;
      }
      if (sameReading(twin, aiOp)) {
        // Оба источника сошлись — самая надёжная из возможных ситуаций.
        out.push({
          op: twin,
          provenance: "algorithm+ai",
          unitIndex: unit.index,
          aiReasoning: response.reasoning,
          dryRun: { ok: run.ok, message: run.message },
        });
        continue;
      }
      // Разошлись. Выбор оставляем человеку, а не большинству голосов:
      // тихо предпочесть любую из версий — значит спрятать разногласие,
      // которое как раз и указывает на сложное место в документе.
      const twinRun = dryRun(twin, offerParts, opts);
      out.push({
        op: twinRun.ok ? twin : aiOp,
        alternative: twinRun.ok ? aiOp : twin,
        provenance: "conflict",
        unitIndex: unit.index,
        aiReasoning: response.reasoning,
        dryRun: twinRun.ok
          ? { ok: twinRun.ok, message: twinRun.message }
          : { ok: run.ok, message: run.message },
      });
    }

    // Алгоритм нашёл больше правок, чем модель, — лишние не теряем.
    for (let i = accepted.length; i < algorithmOps.length; i++) {
      const op = algorithmOps[i];
      const run = dryRun(op, offerParts, opts);
      out.push({
        op,
        provenance: "algorithm",
        unitIndex: unit.index,
        dryRun: { ok: run.ok, message: run.message },
      });
    }
  }

  settings.onProgress?.(queue.size, queue.size, "Готово");
  return { operations: out, stats: countByProvenance(out), aiErrors, aiUsed };
}

/** Убрать из операции служебные поля перед показом модели. */
function stripInternals(op: Operation) {
  const { id: _id, sourceDoc: _s, rawText: _r, warnings: _w, ...rest } = op;
  return rest;
}

function countByProvenance(ops: HybridOperation[]): Record<Provenance, number> {
  const stats: Record<Provenance, number> = {
    algorithm: 0,
    ai: 0,
    "algorithm+ai": 0,
    conflict: 0,
    manual: 0,
  };
  for (const o of ops) stats[o.provenance] += 1;
  return stats;
}
