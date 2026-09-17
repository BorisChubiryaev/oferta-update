// Человеческие подписи к машинным типам. Вынесены из компонентов, потому что
// нужны и экрану проверки, и отчёту о сборке.
import type { Operation, OpType } from "@/engine/types";
import type { Provenance } from "@/ai/types";

export const OP_TYPE_LABEL: Record<OpType, string> = {
  insert_after: "вставка после слов",
  replace: "изложить в новой редакции",
  replace_footnote: "изложить сноску заново",
  add_footnote: "добавить сноску",
  insert_point: "добавить новый пункт",
  append_table_rows: "добавить строки таблицы",
  replace_table_rows: "изменить строки таблицы",
  sort_table_alpha: "сортировка по алфавиту",
  insert_table_row_alpha: "добавить по алфавиту",
  insert_before: "вставка перед словами",
  replace_sentence: "изложить предложение заново",
  replace_paragraph: "изложить абзац заново",
  append_paragraph: "дополнить абзацем",
  append_sentence: "дополнить предложением",
  replace_words: "заменить слова",
  delete_words: "удалить слова",
  replace_words_global: "заменить по всему тексту",
  delete_paragraph: "исключить абзац",
  delete_footnote: "исключить сноску",
  delete_point: "исключить пункт",
  manual: "ручная обработка",
};

export interface ProvenanceStyle {
  label: string;
  cls: string;
  title: string;
}

export const PROVENANCE: Record<Provenance, ProvenanceStyle> = {
  algorithm: {
    label: "алгоритм",
    cls: "algorithm",
    title: "Правка разобрана формальным правилом. ИИ не привлекался.",
  },
  ai: {
    label: "ИИ",
    cls: "ai",
    title:
      "Алгоритм не разобрал формулировку. Операцию предложила модель, " +
      "и она прошла проверку схемой и сухим прогоном по Оферте.",
  },
  "algorithm+ai": {
    label: "алгоритм + ИИ",
    cls: "both",
    title: "Алгоритм и модель прочитали инструкцию одинаково — наивысшая надёжность.",
  },
  conflict: {
    label: "расхождение",
    cls: "conflict",
    title:
      "Алгоритм и модель прочитали инструкцию по-разному. " +
      "Выбор оставлен человеку: сверьте оба варианта с исходным текстом.",
  },
  manual: {
    label: "вручную",
    cls: "manual",
    title: "Автоматически разобрать не удалось — внесите правку в Word самостоятельно.",
  },
};

export function targetLabel(op: Operation): string {
  const t = op.target;
  switch (t.kind) {
    case "footnote":
      return t.atPoint ? `Сноска к п. ${t.atPoint}` : `Сноска № ${t.number}`;
    case "term":
      return `Термин${t.point ? ` (п. ${t.point})` : ""}: «${t.term}»`;
    case "point":
      return `Пункт ${t.point}${t.heading ? ` — ${t.heading}` : ""}`;
    case "preamble":
      return "Преамбула";
    case "appendix_point":
      return `Приложение № ${t.appendix}, п. ${t.point}`;
    case "appendix_table":
      return `Таблица${t.point ? ` п. ${t.point}` : ""} Приложения № ${t.appendix}`;
  }
}
