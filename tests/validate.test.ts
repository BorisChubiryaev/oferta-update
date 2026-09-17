// Проверка заслона между моделью и документом.
//
// Тесты здесь описывают не «что модель обычно отвечает», а что должно
// произойти, когда она ответит плохо: пропущенный якорь, выдуманный тип,
// заменa текста на самоё себя. Каждый такой случай обязан быть отклонён —
// иначе он окажется в тексте договора.
import { test } from "node:test";
import assert from "node:assert/strict";
import { validateOperation, parseAiResponse, stripOuterQuotes } from "../src/ai/validate.ts";

test("принимает корректную операцию", () => {
  const v = validateOperation({
    type: "insert_after",
    target: { kind: "point", point: "7.4" },
    anchor: "статус наличия расчета кредитного потенциала,",
    payload: "дату истечения срока актуальности расчета кредитного потенциала;",
  });
  assert.equal(v.ok, true);
  assert.equal(v.draft?.type, "insert_after");
  assert.equal(v.draft?.target.kind, "point");
});

test("отклоняет неизвестный тип операции", () => {
  const v = validateOperation({ type: "rewrite_everything", target: { kind: "preamble" } });
  assert.equal(v.ok, false);
  assert.match(v.issues[0].message, /неизвестный тип/);
});

test("отклоняет вставку без якоря", () => {
  const v = validateOperation({
    type: "insert_after",
    target: { kind: "point", point: "7.4" },
    payload: "текст",
  });
  assert.equal(v.ok, false);
  assert.ok(v.issues.some((i) => i.field === "anchor"));
});

test("отклоняет новую редакцию без текста", () => {
  const v = validateOperation({ type: "replace", target: { kind: "point", point: "2.1" } });
  assert.equal(v.ok, false);
  assert.ok(v.issues.some((i) => i.field === "payload"));
});

test("отклоняет замену слов на самоё себя", () => {
  const v = validateOperation({
    type: "replace_words",
    target: { kind: "point", point: "3.1" },
    find: "Банк",
    payload: "Банк",
  });
  assert.equal(v.ok, false);
});

test("отклоняет цель без номера пункта", () => {
  const v = validateOperation({ type: "replace", target: { kind: "point" }, payload: "x" });
  assert.equal(v.ok, false);
});

test("сноска может быть задана пунктом вместо номера", () => {
  const v = validateOperation({
    type: "replace_footnote",
    target: { kind: "footnote", atPoint: "5.3" },
    payload: "новая редакция сноски",
  });
  assert.equal(v.ok, true);
  assert.equal(v.draft?.target.kind === "footnote" && v.draft.target.atPoint, "5.3");
});

test("таблица без строк не проходит", () => {
  const v = validateOperation({
    type: "append_table_rows",
    target: { kind: "appendix_table", appendix: "2" },
  });
  assert.equal(v.ok, false);
});

test("снимает только внешнюю пару кавычек", () => {
  assert.equal(stripOuterQuotes("«новая редакция»"), "новая редакция");
  // Внутренние кавычки — часть названия, их трогать нельзя.
  assert.equal(stripOuterQuotes("«ООО «Ромашка» и ООО «Василёк»»"), "ООО «Ромашка» и ООО «Василёк»");
  // Строка не обёрнута целиком: первая пара закрывается в середине.
  assert.equal(stripOuterQuotes("«А», «Б»"), "«А», «Б»");
  assert.equal(stripOuterQuotes("без кавычек"), "без кавычек");
});

test("ответ модели без обёртки operations тоже разбирается", () => {
  const r = parseAiResponse({ type: "replace", target: { kind: "preamble" }, payload: "x" });
  assert.equal(r.operations.length, 1);
});

test("мусор вместо ответа не роняет разбор", () => {
  assert.deepEqual(parseAiResponse(null).operations, []);
  assert.deepEqual(parseAiResponse("строка").operations, []);
  assert.deepEqual(parseAiResponse({ operations: "не массив" }).operations, []);
});
