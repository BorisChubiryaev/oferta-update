// Ошибки слоя ИИ. Отдельный модуль, потому что типы ошибок нужны и серверному
// клиенту (он помечен `server-only`), и чистому разбору ответа.

/** Ответ модели пришёл, но это не JSON нужного вида. */
export class AiFormatError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "AiFormatError";
  }
}
