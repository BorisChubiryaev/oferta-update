// Диагностика настройки ИИ.
//
// Два режима:
//   GET /api/ai/health          — быстрый ответ для интерфейса: настроен ли ИИ.
//   GET /api/ai/health?probe=1  — живая проверка ключа у OpenRouter.
//
// Проверка спрашивает /api/v1/key, а не модель: ответ мгновенный, токенов не
// тратит и отвечает ровно на нужный вопрос — принимает ли OpenRouter этот ключ.
// Иначе «ключ не годится» и «бесплатная модель сейчас занята» сливаются в один
// невнятный отказ.
//
// Сам ключ наружу не отдаётся ни в каком виде — только его длина и форма.
import { NextResponse } from "next/server";
import { probeKey } from "@/ai/openrouter";
import { checkApiKey } from "@/ai/key";

export const runtime = "nodejs";
export const maxDuration = 30;

export async function GET(req: Request) {
  const check = checkApiKey(process.env.OPENROUTER_API_KEY);
  const base = {
    // «Настроен» для интерфейса — значит ключ задан И прошёл проверку формы.
    // Пускать оператора в гибридный режим с заведомо негодным ключом незачем:
    // он узнает об этом на середине пачки инструкций.
    configured: check.ok,
    model: process.env.OPENROUTER_MODEL ?? "inclusionai/ling-3.0-flash-vl:free",
    problem: check.problem,
    message: check.message,
    shape: check.shape,
  };

  const url = new URL(req.url);
  if (url.searchParams.get("probe") !== "1") return NextResponse.json(base);

  return NextResponse.json({ ...base, probe: await probeKey() });
}
