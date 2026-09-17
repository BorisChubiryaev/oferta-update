// Браузерный конвейер (algorithm-only, без сервера и сети).
// Оферта + N документов «Изменения» -> два файла: объединённый перечень
// изменений и текст Оферты с выделенными правками.
import { loadDocx } from "./docx";
import { applyOperations, applyOneOp } from "./apply";
import { buildCombinedDocx } from "./combined";
import { parseInstructionsOffline, resetIds } from "./offline";
import { documentBlocks, tables } from "./text";
import type { BuildOptions, BuildResult, Operation } from "./types";

export interface ChangeDocInput {
  name: string;
  data: Uint8Array;
}

/** Разобрать все документы «Изменения» в операции (детерминированный парсер). */
export async function parseAllChangeDocs(
  changeDocs: ChangeDocInput[],
): Promise<{ operations: Operation[] }> {
  resetIds();
  const all: Operation[] = [];
  for (const cd of changeDocs) {
    const parts = await loadDocx(cd.data);
    const blocks = documentBlocks(parts.document);
    const docTables = tables(parts.document);
    all.push(...parseInstructionsOffline(blocks, docTables, cd.name));
  }
  return { operations: all };
}

/** Отсортировать операции по порядку следования в Оферте (по позиции цели). */
export async function orderOperations(
  offerData: Uint8Array,
  operations: Operation[],
  opts: BuildOptions,
): Promise<Operation[]> {
  const offer = await loadDocx(offerData);
  const keyed = operations.map((op) => {
    const state = {
      document: offer.document,
      footnotes: offer.footnotes,
      numbering: offer.numbering,
      styles: offer.styles,
    };
    const r = applyOneOp(op, state, opts);
    return { op, key: r.ok ? r.orderKey : Number.MAX_SAFE_INTEGER };
  });
  keyed.sort((a, b) => a.key - b.key);
  return keyed.map((k) => k.op);
}

/** Собрать оба итоговых файла из подтверждённых операций. */
export async function buildOutputs(
  offerData: Uint8Array,
  operations: Operation[],
  opts: BuildOptions = {},
): Promise<BuildResult> {
  const ordered = await orderOperations(offerData, operations, opts);
  const offer = await loadDocx(offerData);
  const { offerDocx, results } = await applyOperations(offer, ordered, opts);
  const combinedDocx = await buildCombinedDocx(ordered, {}, results);
  return {
    offerDocx,
    combinedDocx,
    results,
    orderedOperationIds: ordered.map((o) => o.id),
  };
}
