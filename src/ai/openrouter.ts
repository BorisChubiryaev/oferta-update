// Клиент OpenRouter. Работает ТОЛЬКО на сервере: ключ не должен попадать
// в браузерный бандл ни при каких обстоятельствах.
import "server-only";
import { checkApiKey } from "./key";
import { explainStatus, isJsonModeRejection } from "./protocol";

const ENDPOINT = "https://openrouter.ai/api/v1/chat/completions";
/** Проверка ключа: бесплатна и не тратит токены — годится для диагностики. */
const KEY_ENDPOINT = "https://openrouter.ai/api/v1/key";
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

/**
 * Ключ, пригодный к отправке.
 *
 * Значение обрезается по краям: пробел или перевод строки, приехавший вместе
 * с ключом при копировании, — самая частая причина «401 Missing Authentication
 * header», и чинить её перевыпуском ключа бессмысленно. Форма проверяется до
 * сети, чтобы ошибка называла причину, а не приходила от OpenRouter одинаковой
 * для всех трёх способов ошибиться.
 */
function apiKey(): string {
  const check = checkApiKey(process.env.OPENROUTER_API_KEY);
  if (!check.ok || !check.value) {
    throw new OpenRouterError(check.message ?? "Ключ OpenRouter недействителен", 401, false);
  }
  return check.value;
}

/** Заголовки запроса. Атрибуция OpenRouter на работу не влияет. */
function headers(): Record<string, string> {
  return {
    Authorization: `Bearer ${apiKey()}`,
    "Content-Type": "application/json",
    "HTTP-Referer": process.env.OPENROUTER_SITE_URL ?? "https://localhost",
    "X-Title": process.env.OPENROUTER_SITE_NAME ?? "genOferta AI",
  };
}

export interface KeyProbe {
  ok: boolean;
  status?: number;
  message: string;
  /** Название ключа в кабинете OpenRouter — подтверждает, что это тот ключ. */
  label?: string;
  /** Остаток лимита, если OpenRouter его сообщает. */
  usage?: number;
  limit?: number | null;
}

/**
 * Живая проверка ключа.
 *
 * Спрашиваем не модель, а /api/v1/key: ответ приходит мгновенно, не тратит
 * токены и отвечает ровно на нужный вопрос — принимает ли OpenRouter этот ключ.
 * Проверять ключ обращением к модели значит путать два разных отказа: «ключ не
 * годится» и «бесплатная модель сейчас занята».
 */
export async function probeKey(): Promise<KeyProbe> {
  const local = checkApiKey(process.env.OPENROUTER_API_KEY);
  if (!local.ok) {
    return { ok: false, message: local.message ?? "Ключ не прошёл проверку формы" };
  }
  try {
    const res = await fetch(KEY_ENDPOINT, { headers: headers() });
    const body = (await res.json().catch(() => ({}))) as {
      data?: { label?: string; usage?: number; limit?: number | null };
      error?: { message?: string };
    };
    if (!res.ok) {
      return {
        ok: false,
        status: res.status,
        message:
          res.status === 401
            ? "OpenRouter не принял ключ (401). Ключ отозван, удалён или скопирован не полностью. " +
              "Создайте новый на https://openrouter.ai/keys и пропишите его в Vercel, затем Redeploy."
            : `OpenRouter вернул ${res.status}: ${body.error?.message ?? "без пояснения"}`,
      };
    }
    return {
      ok: true,
      status: res.status,
      message: "Ключ принят OpenRouter.",
      label: body.data?.label,
      usage: body.data?.usage,
      limit: body.data?.limit ?? null,
    };
  } catch (e) {
    return {
      ok: false,
      message: `Не удалось связаться с OpenRouter: ${e instanceof Error ? e.message : String(e)}`,
    };
  }
}

/**
 * Поддерживает ли выбранная модель режим JSON (`response_format`).
 *
 * Выясняется только опытным путём: OpenRouter маршрутизирует запрос к
 * провайдеру (Novita, Together и т.д.), и поддержка зависит от связки
 * «модель + провайдер», а не от модели как таковой. Поэтому — не таблица
 * известных моделей, которая устареет к следующей неделе, а один отказ,
 * запомненный на время жизни процесса.
 *
 * `null` — ещё не проверяли.
 */
let jsonModeSupported: boolean | null = null;

/** Явное «не просить JSON» для окружения, где это заведомо не работает. */
function jsonModeDisabledByEnv(): boolean {
  return (process.env.OPENROUTER_JSON_MODE ?? "").toLowerCase() === "off";
}

async function once(
  messages: ChatMessage[],
  signal: AbortSignal,
  jsonMode: boolean,
): Promise<ChatResult> {
  const res = await fetch(ENDPOINT, {
    method: "POST",
    signal,
    headers: headers(),
    body: JSON.stringify({
      model: process.env.OPENROUTER_MODEL ?? DEFAULT_MODEL,
      messages,
      // Разбор инструкции — задача на точность, не на фантазию.
      temperature: 0,
      max_tokens: 4096,
      // Просим JSON — но только если эта связка «модель + провайдер» его умеет.
      // Часть провайдеров на неподдерживаемый флаг отвечает не молчаливым
      // игнорированием, а отказом 400, и тогда запрос не проходит вовсе.
      // Формат ответа этим флагом не держится в любом случае: разбор терпим
      // к обёрткам (см. extractJson), а схема всё равно проверяется отдельно.
      ...(jsonMode ? { response_format: { type: "json_object" } } : {}),
    }),
  });

  if (!res.ok) {
    const body = await res.text().catch(() => "");
    if (jsonMode && isJsonModeRejection(res.status, body)) {
      // Не ошибка настройки и не повод беспокоить оператора: просто эта модель
      // не умеет режим JSON. Запоминаем и повторяем без флага — следующие
      // инструкции пойдут сразу правильным путём, лишнего запроса не будет.
      jsonModeSupported = false;
      return once(messages, signal, false);
    }
    // 429 (лимит бесплатной модели) и 5xx имеет смысл повторить, 4xx — нет.
    const retriable = res.status === 429 || res.status >= 500;
    throw new OpenRouterError(explainStatus(res.status, body), res.status, retriable);
  }
  if (jsonMode) jsonModeSupported = true;

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
      // Пока не выяснили обратное — просим JSON: с ним ответ чище.
      const jsonMode = !jsonModeDisabledByEnv() && jsonModeSupported !== false;
      return await once(messages, ctrl.signal, jsonMode);
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

// Разбор ответа модели (json.ts) и разбор ответов протокола (protocol.ts)
// живут отдельно и переэкспортируются здесь, чтобы у вызывающего кода был
// один вход.
export { isJsonModeRejection, explainStatus } from "./protocol";
