// Настроен ли ИИ. Браузеру нужно знать это ДО загрузки документов, чтобы
// честно показать режим работы, а не обещать гибридный разбор и упереться
// в отсутствующий ключ на середине пачки инструкций.
import { NextResponse } from "next/server";

export const runtime = "nodejs";

export async function GET() {
  return NextResponse.json({
    configured: Boolean(process.env.OPENROUTER_API_KEY),
    model: process.env.OPENROUTER_MODEL ?? "inclusionai/ling-3.0-flash-vl:free",
  });
}
