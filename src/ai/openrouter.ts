// Клиент OpenRouter. Работает ТОЛЬКО на сервере: ключ не должен попадать
// в браузерный бандл ни при каких обстоятельствах.
import "server-only";

const ENDPOINT = "https://openrouter.ai/api/v1/chat/completions";
const DEFAULT_MODEL = "inclusionai/ling-3.0-flash-vl:free";

/** Бесплатные модели OpenRouter отвечают неровно — отсюда таймаут и повторы. */
const TIMEOUT_MS = 60_000;
const RETRIES = 2;

export class OpenRouterError extends Error {
  constructor(
    message: string,
    readonly status?: number,
    readonly retriable = false,
  ) {
    super(message);
    this.name = "OpenRouterError";
  }
}

export interface ChatMessage {
  role: "system" | "user" | "assistant";
  content: string;
}

export interface ChatResult {
  text: string;
  model: string;
}

function apiKey(): string {
  const key = process.env.OPENROUTER_API_KEY;
  if (!key) {
    throw new OpenRouterError(
      "OPENROUTER_API_KEY не задан. Добавьте ключ в переменные окружения Vercel " +
        "(Settings → Environment Variables) или в .env.local для локального запуска.",
      undefined,
      false,
    );
  }
  return key;
}

async function once(messages: ChatMessage[], signal: AbortSignal): Promise<ChatResult> {
  const res = await fetch(ENDPOINT, {
    method: "POST",
    signal,
    headers: {
      Authorization: `Bearer ${apiKey()}`,
      "Content-Type": "application/json",
      // OpenRouter просит атрибуцию; на работу запроса она не влияет.
      "HTTP-Referer": process.env.OPENROUTER_SITE_URL ?? "https://localhost",
      "X-Title": process.env.OPENROUTER_SITE_NAME ?? "genOferta AI",
    },
    body: JSON.stringify({
      model: process.env.OPENROUTER_MODEL ?? DEFAULT_MODEL,
      messages,
      // Разбор инструкции — задача на точность, не на фантазию.
      temperature: 0,
      max_tokens: 4096,
      // Просим JSON. Не все модели уважают этот флаг, поэтому ответ всё равно
      // проходит через терпимый к обёрткам разбор (см. extractJson).
      response_format: { type: "json_object" },
    }),
  });

  if (!res.ok) {
    const body = await res.text().catch(() => "");
    // 429 (лимит бесплатной модели) и 5xx имеет смысл повторить, 4xx — нет.
    const retriable = res.status === 429 || res.status >= 500;
    throw new OpenRouterError(
      `OpenRouter вернул ${res.status}: ${body.slice(0, 300)}`,
      res.status,
      retriable,
    );
  }

  const data = (await res.json()) as {
    choices?: { message?: { content?: string } }[];
    model?: string;
    error?: { message?: string };
  };
  if (data.error?.message) throw new OpenRouterError(data.error.message, 200, true);
  const text = data.choices?.[0]?.message?.content;
  if (!text) throw new OpenRouterError("Пустой ответ модели", 200, true);
  return { text, model: data.model ?? DEFAULT_MODEL };
}

/** Запрос к модели с таймаутом и повторами по возвратным ошибкам. */
export async function chat(messages: ChatMessage[]): Promise<ChatResult> {
  let last: unknown;
  for (let attempt = 0; attempt <= RETRIES; attempt++) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), TIMEOUT_MS);
    try {
      return await once(messages, ctrl.signal);
    } catch (e) {
      last = e;
      const retriable =
        (e instanceof OpenRouterError && e.retriable) ||
        (e instanceof Error && e.name === "AbortError");
      if (!retriable || attempt === RETRIES) break;
      // Пауза растёт: бесплатные модели отвечают 429 пачками.
      await new Promise((r) => setTimeout(r, 800 * 2 ** attempt));
    } finally {
      clearTimeout(timer);
    }
  }
  throw last instanceof Error ? last : new OpenRouterError(String(last));
}

// Разбор ответа модели живёт отдельно (см. json.ts) и переэкспортируется
// здесь, чтобы у вызывающего кода был один вход.
export { extractJson } from "./json";
