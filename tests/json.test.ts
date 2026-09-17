// Разбор ответа модели. Модели оборачивают JSON во что угодно, и падать
// из-за обёртки, имея на руках валидный JSON, — значит терять правку.
import { test } from "node:test";
import assert from "node:assert/strict";
import { extractJson } from "../src/ai/json.ts";

test("чистый JSON", () => {
  assert.deepEqual(extractJson('{"a":1}'), { a: 1 });
});

test("JSON в ограде ```json", () => {
  assert.deepEqual(extractJson('Вот результат:\n```json\n{"a":1}\n```\n'), { a: 1 });
});

test("JSON с текстом вокруг", () => {
  assert.deepEqual(extractJson('Я разобрал инструкцию. {"a":[1,2]} Надеюсь, помог.'), {
    a: [1, 2],
  });
});

test("фигурные скобки внутри строки не сбивают границы объекта", () => {
  const raw = 'болтовня {"payload":"текст со скобкой } и кавычкой \\" внутри","b":2} хвост';
  assert.deepEqual(extractJson(raw), { payload: 'текст со скобкой } и кавычкой " внутри', b: 2 });
});

test("не-JSON приводит к ошибке, а не к тихому undefined", () => {
  assert.throws(() => extractJson("модель отказалась отвечать"));
});
