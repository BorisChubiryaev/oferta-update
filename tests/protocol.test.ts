// Разбор ответов OpenRouter на уровне протокола.
//
// Главное здесь — распознавание отказа «модель не умеет structured-outputs».
// OpenRouter маршрутизирует запрос к провайдеру (Novita, Together и др.), и
// поддержка режима JSON зависит от связки «модель + провайдер». Провайдер
// заворачивает свой текст в metadata.raw с экранированными кавычками, так что
// проверка обязана работать по телу ответа целиком, а не по разобранному полю.
import { test } from "node:test";
import assert from "node:assert/strict";
import { isJsonModeRejection, explainStatus } from "../src/ai/protocol.ts";

// Настоящий ответ OpenRouter на inclusionai/ling-3.0-flash-vl через Novita.
const REAL_REJECTION =
  '{"error":{"message":"Provider returned error","code":400,"metadata":{"raw":' +
  '"{\\"code\\":400,\\"reason\\":\\"INVALID_REQUEST_BODY\\",\\"message\\":\\"model: ' +
  'inclusionai/ling-3.0-flash-vl does not support feature: structured-outputs\\"}",' +
  '"provider_name":"Novita"}}}';

test("настоящий отказ провайдера опознаётся", () => {
  assert.equal(isJsonModeRejection(400, REAL_REJECTION), true);
});

test("другие формулировки того же отказа тоже опознаются", () => {
  for (const body of [
    '{"error":{"message":"response_format is not supported"}}',
    '{"error":{"message":"json_object unsupported for this model"}}',
    '{"error":{"message":"structured outputs not available"}}',
    '{"error":{"message":"json_schema is not supported by provider"}}',
  ]) {
    assert.equal(isJsonModeRejection(400, body), true, body);
  }
});

test("посторонний 400 не принимается за отказ от режима JSON", () => {
  // Иначе запрос молча повторится без флага и упадёт второй раз — с тем же
  // результатом, но вдвое дольше и с запутанной диагностикой.
  assert.equal(isJsonModeRejection(400, '{"error":{"message":"messages: too long"}}'), false);
});

test("не-400 отказом от режима JSON не считается", () => {
  // 429 c упоминанием json_object — это лимит, а не отсутствие возможности.
  assert.equal(isJsonModeRejection(429, '{"error":"json_object rate limited"}'), false);
  assert.equal(isJsonModeRejection(500, REAL_REJECTION), false);
});

test("коды ответов объясняются по существу", () => {
  assert.match(explainStatus(401, ""), /ключ/i);
  assert.match(explainStatus(402, ""), /лимит|средств/i);
  assert.match(explainStatus(404, ""), /OPENROUTER_MODEL/);
  assert.match(explainStatus(429, ""), /429/);
  assert.match(explainStatus(400, ""), /400/);
  // Неизвестный код не теряет тело ответа — по нему и разбираются.
  assert.match(explainStatus(418, "я чайник"), /418.*я чайник/s);
});
