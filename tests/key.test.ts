// Проверка ключа до обращения к сети.
//
// Каждый тест здесь — реальный способ ошибиться при настройке Vercel. Все они
// приводят к одному и тому же «401 Missing Authentication header» от
// OpenRouter, по которому непонятно, что чинить; задача проверки — различить
// их и назвать причину.
import { test } from "node:test";
import assert from "node:assert/strict";
import { checkApiKey } from "../src/ai/key.ts";

const VALID = "sk-or-v1-" + "a".repeat(64);

test("настоящий ключ OpenRouter проходит", () => {
  const r = checkApiKey(VALID);
  assert.equal(r.ok, true);
  assert.equal(r.value, VALID);
});

test("пробел и перевод строки по краям срезаются, а не ломают запрос", () => {
  // Самая частая причина 401: ключ скопирован вместе с пробелом.
  const r = checkApiKey(`  ${VALID}\n`);
  assert.equal(r.ok, true);
  assert.equal(r.value, VALID);
  assert.equal(r.shape.hadWhitespace, true);
});

test("пустая и незаданная переменная различимы по сообщению", () => {
  for (const v of [undefined, "", "   "]) {
    const r = checkApiKey(v);
    assert.equal(r.ok, false);
    assert.equal(r.problem, "missing");
  }
});

test("ключ чужого сервиса отклоняется до сети", () => {
  // «sk-» + hex — формат других OpenAI-совместимых шлюзов, OpenRouter его
  // не примет, и узнать об этом лучше на своей стороне.
  const r = checkApiKey("sk-" + "b3e992d3174c4bacb1fea41dbf5bd09b".repeat(2));
  assert.equal(r.ok, false);
  assert.equal(r.problem, "wrong-format");
  assert.match(r.message ?? "", /sk-or-v1-/);
});

test("название модели в переменной ключа опознаётся отдельно", () => {
  const r = checkApiKey("inclusionai/ling-3.0-flash-vl:free");
  assert.equal(r.ok, false);
  assert.equal(r.problem, "looks-like-model");
  assert.match(r.message ?? "", /OPENROUTER_MODEL/);
});

test("кавычки вокруг значения опознаются отдельно", () => {
  const r = checkApiKey(`"${VALID}"`);
  assert.equal(r.ok, false);
  assert.equal(r.problem, "quoted");
});

test("характеристика значения не содержит самого ключа", () => {
  const r = checkApiKey(VALID);
  const dump = JSON.stringify(r.shape);
  assert.ok(!dump.includes(VALID), "ключ не должен попадать в диагностику");
  assert.equal(r.shape.length, VALID.length);
});
