// Конвейер целиком — на реальной Оферте, но с подставной моделью.
//
// Настоящий OpenRouter здесь не нужен и вреден: тест должен проверять НАШИ
// гарантии, а не сегодняшнее настроение бесплатной модели. Поэтому вместо
// сети — функция, которая отвечает ровно тем, что нужно проверить: выдумкой,
// мусором, отказом или верным разбором.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { runHybrid } from "../src/ai/hybrid.ts";
import type { AiCaller } from "../src/ai/hybrid.ts";

const OFFER = "./samples/Версия 70.docx";
const CHANGE = "./samples/Изменения_4218-1_71.docx";

// Реальные документы в репозиторий не кладутся (внутренние документы банка),
// поэтому без них тест пропускается, а не падает.
const haveSamples = existsSync(OFFER) && existsSync(CHANGE);

async function input() {
  return {
    offer: new Uint8Array(await readFile(OFFER)),
    changeDocs: [{ name: "Изменения_71.docx", data: new Uint8Array(await readFile(CHANGE)) }],
  };
}

const neverCalled: AiCaller = async () => {
  throw new Error("модель не должна вызываться в этом режиме");
};

test("без ИИ работает прежний алгоритмический разбор", { skip: !haveSamples }, async () => {
  const r = await runHybrid(await input(), { useAi: false, verifyConfident: false }, neverCalled);
  assert.equal(r.aiUsed, false);
  assert.ok(r.operations.length > 0, "правки должны быть найдены");
  assert.ok(
    r.operations.every((o) => o.provenance === "algorithm" || o.provenance === "manual"),
    "без ИИ других источников быть не может",
  );
});

test("выдумка модели отсекается сухим прогоном", { skip: !haveSamples }, async () => {
  // Пункта 99.99 в Оферте нет. Схему такая операция проходит — ловит её только
  // прогон по реальному документу.
  const liar: AiCaller = async () => ({
    operations: [
      {
        type: "replace",
        target: { kind: "point", point: "99.99" },
        payload: "выдуманная редакция несуществующего пункта",
      },
    ],
    reasoning: "выдумка",
  });
  const r = await runHybrid(await input(), { useAi: true, verifyConfident: true }, liar);
  const fabricated = r.operations.filter(
    (o) => o.op.target.kind === "point" && o.op.target.point === "99.99",
  );
  assert.equal(fabricated.length, 0, "несуществующая цель не должна попасть в результат");
});

test("мусор от модели не роняет сборку", { skip: !haveSamples }, async () => {
  const garbage: AiCaller = async () => ({
    operations: [{ nonsense: true }, "строка", null, 42],
    reasoning: "мусор",
  });
  const r = await runHybrid(await input(), { useAi: true, verifyConfident: true }, garbage);
  assert.ok(r.operations.length > 0, "результат алгоритма должен сохраниться");
  assert.ok(
    r.operations.every((o) => typeof o.op.type === "string"),
    "все операции остаются структурно корректными",
  );
});

test("падение ИИ не теряет правки алгоритма", { skip: !haveSamples }, async () => {
  const broken: AiCaller = async () => {
    throw new Error("OpenRouter вернул 429");
  };
  const withAi = await runHybrid(await input(), { useAi: true, verifyConfident: true }, broken);
  const withoutAi = await runHybrid(
    await input(),
    { useAi: false, verifyConfident: false },
    neverCalled,
  );
  assert.ok(withAi.aiErrors.length > 0, "ошибка ИИ должна быть показана оператору");
  assert.equal(
    withAi.operations.length,
    withoutAi.operations.length,
    "при недоступном ИИ результат не должен становиться беднее алгоритмического",
  );
});

test("совпавшее прочтение помечается как алгоритм+ИИ", { skip: !haveSamples }, async () => {
  // Модель возвращает ровно черновик алгоритма — так ведёт себя режим verify,
  // когда алгоритм прав.
  const echo: AiCaller = async (req) => ({
    operations: Array.isArray(req.algorithmDraft) ? (req.algorithmDraft as unknown[]) : [],
    reasoning: "черновик верен",
  });
  const r = await runHybrid(await input(), { useAi: true, verifyConfident: true }, echo);
  assert.ok(
    r.operations.some((o) => o.provenance === "algorithm+ai"),
    "согласие двух источников должно быть видно оператору",
  );
});

test("модель разбирает то, что алгоритм пометил как ручное", { skip: !haveSamples }, async () => {
  // Инструкция 71/[1] — вставка по якорю, который алгоритм не нашёл дословно.
  // Подставная модель даёт якорь, который в Оферте действительно есть.
  const helper: AiCaller = async (req) => {
    if (req.mode !== "parse") return { operations: [] };
    return {
      operations: [
        {
          type: "replace_words_global",
          target: { kind: "point", point: "7.4" },
          find: "кредитного потенциала",
          payload: "кредитного потенциала клиента",
        },
      ],
      reasoning: "разобрано моделью",
    };
  };
  const r = await runHybrid(await input(), { useAi: true, verifyConfident: false }, helper);
  // Проверяем не конкретную правку, а инвариант: всё, что пришло от модели,
  // применимо к документу — иначе оно сюда бы не попало.
  for (const o of r.operations.filter((x) => x.provenance === "ai")) {
    assert.equal(o.dryRun?.ok, true, `операция от ИИ обязана применяться: ${o.dryRun?.message}`);
  }
});
