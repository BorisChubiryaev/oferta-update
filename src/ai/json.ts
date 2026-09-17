// Разбор JSON из ответа модели.
//
// Вынесено из клиента OpenRouter отдельным модулем сознательно: клиент помечен
// `server-only` (он трогает ключ), а эта функция — чистая, и её нужно уметь
// проверять тестами, не поднимая окружение Next.
import { AiFormatError } from "./errors";

/**
 * Достать JSON из ответа модели.
 *
 * Даже с response_format модели оборачивают ответ в ```json … ``` или
 * предваряют его фразой. Падать из-за обёртки, имея на руках валидный JSON,
 * — значит терять правку на ровном месте, поэтому разбор терпимый: сначала
 * как есть, потом без ограды, потом первый сбалансированный объект.
 */
export function extractJson(raw: string): unknown {
  const attempts: string[] = [raw.trim()];
  const fence = raw.match(/```(?:json)?\s*([\s\S]*?)```/);
  if (fence) attempts.push(fence[1].trim());
  const brace = balancedObject(raw);
  if (brace) attempts.push(brace);
  for (const a of attempts) {
    try {
      return JSON.parse(a);
    } catch {
      /* пробуем следующий вариант */
    }
  }
  throw new AiFormatError(`Ответ модели не является JSON: ${raw.slice(0, 200)}`);
}

/** Первый сбалансированный {...}, с учётом скобок внутри строк. */
function balancedObject(s: string): string | null {
  const start = s.indexOf("{");
  if (start < 0) return null;
  let depth = 0;
  let inStr = false;
  let esc = false;
  for (let i = start; i < s.length; i++) {
    const ch = s[i];
    if (inStr) {
      if (esc) esc = false;
      else if (ch === "\\") esc = true;
      else if (ch === '"') inStr = false;
      continue;
    }
    if (ch === '"') inStr = true;
    else if (ch === "{") depth++;
    else if (ch === "}") {
      depth--;
      if (depth === 0) return s.slice(start, i + 1);
    }
  }
  return null;
}
