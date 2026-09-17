// Прогон конвейера по реальным документам — без сети и без ИИ.
//
// Проверяет ровно то, что ломается молча: сколько правок распознано, сколько
// из них движок реально применяет к Оферте и какие остаются оператору. Числа
// печатаются, а не сверяются с эталоном: эталон здесь быстро устаревает, а
// глазами оператора отчёт читается за секунду.
//
// Запуск: npm run regress -- ./samples
import { readFile, readdir, writeFile, mkdir } from "node:fs/promises";
import { join } from "node:path";
import { loadDocx } from "../src/engine/docx.ts";
import { analyzeChangeDoc } from "../src/engine/offline.ts";
import { documentBlocks, tables } from "../src/engine/text.ts";
import { buildOfferDigest, digestToPrompt } from "../src/engine/offer-digest.ts";
import { applyOneOp } from "../src/engine/apply.ts";
import { buildOutputs } from "../src/engine/pipeline.ts";
import type { Operation } from "../src/engine/types.ts";

const dir = process.argv[2] ?? "./samples";

function isOffer(name: string): boolean {
  return /Версия/i.test(name);
}

const main = async () => {
  const files = (await readdir(dir)).filter((f) => f.endsWith(".docx") && !f.startsWith("~"));
  const offerName = files.find(isOffer);
  if (!offerName) throw new Error(`В ${dir} нет файла Оферты (имя должно содержать «Версия»)`);

  const offerData = new Uint8Array(await readFile(join(dir, offerName)));
  const offer = await loadDocx(offerData);
  const digest = buildOfferDigest(offer);

  console.log(`ОФЕРТА: ${offerName}`);
  console.log(`  пунктов: ${digest.points.length}`);
  console.log(`  разделов: ${digest.sections.length}`);
  console.log(`  сносок: ${digest.footnoteCount}`);
  console.log(`  терминов: ${digest.terms.length}`);
  console.log(`  приложений: ${digest.appendices.join(", ") || "—"}`);
  const prompt = digestToPrompt(digest);
  console.log(`  размер карты для ИИ: ${prompt.length} символов (~${Math.round(prompt.length / 3)} токенов)`);
  console.log();

  const all: Operation[] = [];
  let totalUnits = 0;
  let applied = 0;
  let failed = 0;
  let manual = 0;

  for (const name of files.filter((f) => !isOffer(f))) {
    const parts = await loadDocx(new Uint8Array(await readFile(join(dir, name))));
    const { units } = analyzeChangeDoc(documentBlocks(parts.document), tables(parts.document), name);
    totalUnits += units.length;
    console.log(`ИЗМЕНЕНИЯ: ${name} — инструкций: ${units.length}`);

    for (const u of units) {
      for (const op of u.operations) {
        all.push(op);
        if (op.type === "manual") {
          manual += 1;
          console.log(`  ✋ [${u.index}] ${op.note ?? "нужен человек"}`);
          console.log(`       «${u.text.slice(0, 140)}…»`);
          continue;
        }
        // Сухой прогон по свежей копии строк: applyOneOp пишет в переданный
        // объект, но строки неизменяемы, поэтому исходная Оферта не портится.
        const r = applyOneOp(
          op,
          {
            document: offer.document,
            footnotes: offer.footnotes,
            numbering: offer.numbering,
            styles: offer.styles,
          },
          {},
        );
        if (r.ok) {
          applied += 1;
          console.log(`  ✅ [${u.index}] ${op.type} → ${r.message}`);
        } else {
          failed += 1;
          console.log(`  ⚠️  [${u.index}] ${op.type} → ${r.message}`);
          console.log(`       «${u.text.slice(0, 140)}…»`);
        }
      }
    }
    console.log();
  }

  console.log("ИТОГО");
  console.log(`  инструкций разобрано: ${totalUnits}`);
  console.log(`  операций получено:    ${all.length}`);
  console.log(`  применяется движком:  ${applied}`);
  console.log(`  не применяется:       ${failed}   ← кандидаты на разбор через ИИ`);
  console.log(`  требует человека:     ${manual}   ← кандидаты на разбор через ИИ`);

  // Полная сборка: проверяем, что на выходе получаются валидные docx.
  const usable = all.filter((o) => o.type !== "manual");
  const out = await buildOutputs(offerData, usable, { highlightMode: "color", showOld: true });
  await mkdir("./out", { recursive: true });
  await writeFile("./out/offer.docx", out.offerDocx);
  await writeFile("./out/combined.docx", out.combinedDocx);
  console.log();
  console.log(`  СБОРКА: out/offer.docx (${Math.round(out.offerDocx.length / 1024)} КБ), ` +
    `out/combined.docx (${Math.round(out.combinedDocx.length / 1024)} КБ)`);
  console.log(`  успешных операций при сборке: ${out.results.filter((r) => r.ok).length}/${out.results.length}`);
};

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
