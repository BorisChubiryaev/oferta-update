// Прокси к OpenRouter. Существует ровно по одной причине: ключ не должен
// попадать в браузер. Никакой логики разбора здесь нет — она в src/ai.
import { NextResponse } from "next/server";
import { chat, extractJson, OpenRouterError } from "@/ai/openrouter";
import { SYSTEM_PROMPT, buildUserPrompt } from "@/ai/prompts";
import { parseAiResponse } from "@/ai/validate";
import type { AiParseRequest } from "@/ai/types";

export const runtime = "nodejs";
// Бесплатная модель отвечает небыстро; дефолтных 10 секунд Vercel не хватает.
export const maxDuration = 60;

/** Инструкция длиннее этого — почти наверняка склеенный мусор, а не правка. */
const MAX_INSTRUCTION = 20_000;

export async function POST(req: Request) {
  let body: AiParseRequest;
  try {
    body = (await req.json()) as AiParseRequest;
  } catch {
    return NextResponse.json({ error: "Некорректный JSON запроса" }, { status: 400 });
  }

  if (!body?.instruction || typeof body.instruction !== "string") {
    return NextResponse.json({ error: "Не передан текст инструкции" }, { status: 400 });
  }
  if (body.instruction.length > MAX_INSTRUCTION) {
    return NextResponse.json(
      { error: `Инструкция длиннее ${MAX_INSTRUCTION} символов` },
      { status: 413 },
    );
  }

  try {
    const result = await chat([
      { role: "system", content: SYSTEM_PROMPT },
      {
        role: "user",
        content: buildUserPrompt({
          instruction: body.instruction,
          tables: Array.isArray(body.tables) ? body.tables : undefined,
          offerDigest: typeof body.offerDigest === "string" ? body.offerDigest : "",
          algorithmDraft: body.algorithmDraft,
          mode: body.mode === "verify" ? "verify" : "parse",
        }),
      },
    ]);
    const parsed = parseAiResponse(extractJson(result.text));
    return NextResponse.json({ ...parsed, model: result.model });
  } catch (e) {
    const status = e instanceof OpenRouterError ? (e.status ?? 502) : 502;
    return NextResponse.json(
      { error: e instanceof Error ? e.message : "Ошибка обращения к модели" },
      // 401/403 (плохой ключ) отдаём как есть — оператору важно отличить
      // «модель не справилась» от «ключ не настроен».
      { status: status === 401 || status === 403 ? status : 502 },
    );
  }
}
