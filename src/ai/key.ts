// Проверка ключа OpenRouter ДО обращения к сети.
//
// Сам по себе ключ — это строка из переменной окружения, и ошибиться в ней
// можно ровно тремя способами: не задать, задать не тот, задать с мусором по
// краям (пробел или перевод строки при копировании — самая частая причина).
// Все три OpenRouter возвращает одинаково невнятным «401 Missing
// Authentication header», по которому невозможно понять, что чинить.
// Поэтому форма ключа проверяется здесь, на своей стороне, и ошибка называет
// причину прямо.
//
// Модуль намеренно не помечен `server-only`: он ничего не читает из окружения
// сам — значение ему передают. Это позволяет проверять его тестами.

/**
 * Формат ключа OpenRouter: `sk-or-v1-` и дальше буквы с цифрами.
 * Ключи других сервисов (`sk-` + hex у OpenAI-совместимых шлюзов) сюда не
 * подходят — и это стоит поймать до сети, а не после.
 */
const OPENROUTER_KEY_RE = /^sk-or-v1-[A-Za-z0-9]+$/;

export type KeyProblem =
  | "missing"
  | "whitespace"
  | "quoted"
  | "wrong-format"
  | "looks-like-model";

export interface KeyCheck {
  ok: boolean;
  problem?: KeyProblem;
  /** Готовое объяснение для оператора — без самого ключа. */
  message?: string;
  /** Ключ, пригодный к отправке (обрезанный по краям). */
  value?: string;
  /**
   * Безопасная характеристика значения: длина и форма, но НЕ содержимое.
   * Нужна, чтобы диагностировать проблему по экрану health, не пересылая ключ.
   */
  shape: {
    length: number;
    hadWhitespace: boolean;
    startsWithSkOrV1: boolean;
  };
}

const VERCEL_HINT =
  "В Vercel: Settings → Environment Variables → OPENROUTER_API_KEY. " +
  "После изменения переменной нужен Redeploy — на лету она не подхватывается.";

export function checkApiKey(raw: string | undefined): KeyCheck {
  const value = (raw ?? "").trim();
  const shape = {
    length: value.length,
    hadWhitespace: raw !== undefined && raw !== value,
    startsWithSkOrV1: value.startsWith("sk-or-v1-"),
  };

  if (!value) {
    return {
      ok: false,
      problem: "missing",
      shape,
      message: `OPENROUTER_API_KEY не задан (или пуст). ${VERCEL_HINT}`,
    };
  }

  // Кавычки вокруг значения: в вебе их не ставят, но при копировании из .env
  // или из примера кода они переезжают вместе со строкой.
  if (/^["'].*["']$/.test(value)) {
    return {
      ok: false,
      problem: "quoted",
      shape,
      message:
        "Значение OPENROUTER_API_KEY взято в кавычки. В переменные Vercel ключ " +
        `вписывается без кавычек. ${VERCEL_HINT}`,
    };
  }

  // Классическая путаница: в переменную с ключом попало название модели.
  if (value.includes("/") || value.includes(":")) {
    return {
      ok: false,
      problem: "looks-like-model",
      shape,
      message:
        "В OPENROUTER_API_KEY лежит не ключ, а похоже на название модели " +
        "(есть «/» или «:»). Название модели задаётся отдельной переменной " +
        `OPENROUTER_MODEL, а ключ — в OPENROUTER_API_KEY. ${VERCEL_HINT}`,
    };
  }

  if (!OPENROUTER_KEY_RE.test(value)) {
    return {
      ok: false,
      problem: "wrong-format",
      shape,
      message:
        "OPENROUTER_API_KEY не похож на ключ OpenRouter: они начинаются с «sk-or-v1-». " +
        "Ключ вида «sk-» + шестнадцатеричная строка — это ключ другого сервиса, " +
        "OpenRouter его не примет. Создайте ключ на https://openrouter.ai/keys " +
        `и скопируйте целиком. ${VERCEL_HINT}`,
    };
  }

  return { ok: true, value, shape };
}
