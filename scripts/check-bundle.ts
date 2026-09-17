// Не утёк ли ключ в браузер.
//
// `server-only` ловит неверный импорт на сборке, но ключ можно занести в
// клиент и мимо него — например, переименовав переменную в NEXT_PUBLIC_*.
// Проверка дешёвая, а цена ошибки — опубликованный на весь интернет ключ,
// поэтому она делается по факту сборки, а не по намерениям кода.
//
// Запуск: npm run check:bundle (после npm run build)
import { readdir, readFile } from "node:fs/promises";
import { join } from "node:path";

const CHUNKS = ".next/static";

/** Что искать в клиентских чанках. */
const FORBIDDEN: { name: string; re: RegExp }[] = [
  // Ключ OpenRouter в любом из его форматов.
  { name: "ключ OpenRouter", re: /\bsk-(?:or-v1-)?[a-f0-9]{32,}\b/ },
  // Прямое обращение к OpenRouter из браузера: даже без ключа это значит, что
  // запрос идёт мимо серверного маршрута.
  { name: "адрес OpenRouter", re: /openrouter\.ai\/api/ },
  // Значение ключа, вынесенное в публичную переменную.
  { name: "ключ в NEXT_PUBLIC_*", re: /NEXT_PUBLIC_[A-Z_]*(?:KEY|TOKEN|SECRET)/ },
];

async function* walk(dir: string): AsyncGenerator<string> {
  for (const e of await readdir(dir, { withFileTypes: true })) {
    const p = join(dir, e.name);
    if (e.isDirectory()) yield* walk(p);
    else if (e.name.endsWith(".js")) yield p;
  }
}

const main = async () => {
  const hits: string[] = [];
  let checked = 0;
  for await (const file of walk(CHUNKS)) {
    checked += 1;
    const text = await readFile(file, "utf8");
    for (const f of FORBIDDEN) {
      const m = text.match(f.re);
      if (m) hits.push(`${file}: ${f.name} → ${m[0].slice(0, 24)}…`);
    }
  }
  if (hits.length) {
    console.error(`❌ В клиентском бандле найдено то, чего там быть не должно:`);
    for (const h of hits) console.error(`   ${h}`);
    process.exit(1);
  }
  console.log(`✅ Клиентский бандл чист (проверено файлов: ${checked}).`);
};

main().catch((e) => {
  console.error(`Проверка не выполнена: ${e instanceof Error ? e.message : e}`);
  console.error("Сначала выполните npm run build.");
  process.exit(1);
});
