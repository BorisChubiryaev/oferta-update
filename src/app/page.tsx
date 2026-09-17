"use client";

import { useEffect, useMemo, useState } from "react";
import FileDrop, { type PickedFile } from "@/components/FileDrop";
import OperationCard from "@/components/OperationCard";
import { runHybrid } from "@/ai/hybrid";
import { buildOutputs } from "@/engine/pipeline";
import { PROVENANCE } from "@/lib/labels";
import type { HybridOperation, HybridResult, Provenance } from "@/ai/types";
import type { ApplyResult, HighlightMode, Operation } from "@/engine/types";

type Stage = "upload" | "review" | "done";

interface AiHealth {
  configured: boolean;
  model: string;
  /** Что именно не так с ключом — заполняется, когда configured=false. */
  message?: string;
  problem?: string;
}

interface KeyProbe {
  ok: boolean;
  message: string;
  label?: string;
}

function download(data: Uint8Array, name: string) {
  // Копия в собственный ArrayBuffer: у Uint8Array из JSZip буфер может быть
  // общим и длиннее самого массива, и Blob тогда прихватит чужие байты.
  const copy = new Uint8Array(data.length);
  copy.set(data);
  const url = URL.createObjectURL(
    new Blob([copy], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }),
  );
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}

export default function Page() {
  const [stage, setStage] = useState<Stage>("upload");
  const [offer, setOffer] = useState<PickedFile[]>([]);
  const [changes, setChanges] = useState<PickedFile[]>([]);

  const [useAi, setUseAi] = useState(true);
  const [verifyConfident, setVerifyConfident] = useState(false);
  const [highlightMode, setHighlightMode] = useState<HighlightMode>("color");
  const [showOld, setShowOld] = useState(true);

  const [health, setHealth] = useState<AiHealth | null>(null);
  const [probe, setProbe] = useState<KeyProbe | null>(null);
  const [probing, setProbing] = useState(false);
  const [busy, setBusy] = useState<string | null>(null);
  const [progress, setProgress] = useState({ done: 0, total: 0 });
  const [error, setError] = useState<string | null>(null);

  const [result, setResult] = useState<HybridResult | null>(null);
  const [items, setItems] = useState<HybridOperation[]>([]);
  const [disabled, setDisabled] = useState<Set<string>>(new Set());
  const [applyReport, setApplyReport] = useState<ApplyResult[] | null>(null);

  useEffect(() => {
    fetch("/api/ai/health")
      .then((r) => r.json())
      .then((h: AiHealth) => {
        setHealth(h);
        // Обещать гибридный разбор без ключа нечестно — сразу переключаемся
        // в алгоритмический режим и говорим об этом.
        if (!h.configured) setUseAi(false);
      })
      .catch(() => setHealth({ configured: false, model: "" }));
  }, []);

  const analyze = async () => {
    if (!offer.length || !changes.length) return;
    setError(null);
    setBusy("Разбор документов");
    setProgress({ done: 0, total: 0 });
    try {
      const res = await runHybrid(
        { offer: offer[0].data, changeDocs: changes },
        {
          useAi: useAi && Boolean(health?.configured),
          verifyConfident,
          onProgress: (done, total, stage) => {
            setProgress({ done, total });
            setBusy(stage);
          },
        },
      );
      setResult(res);
      setItems(res.operations);
      // Правки, которые движок не применит, по умолчанию выключены: включённая
      // и молча провалившаяся операция создаёт ложное впечатление, что всё учтено.
      setDisabled(
        new Set(
          res.operations
            .filter((o) => o.op.type === "manual" || o.dryRun?.ok === false)
            .map((o) => o.op.id),
        ),
      );
      setStage("review");
    } catch (e) {
      setError(e instanceof Error ? e.message : "Не удалось разобрать документы");
    } finally {
      setBusy(null);
    }
  };

  const build = async () => {
    if (!offer.length) return;
    setError(null);
    setBusy("Сборка документов");
    try {
      const selected = items
        .filter((i) => !disabled.has(i.op.id) && i.op.type !== "manual")
        .map((i) => i.op);
      const out = await buildOutputs(offer[0].data, selected, {
        highlightMode,
        showOld,
        author: "genOferta AI",
      });
      const stamp = new Date().toISOString().slice(0, 10);
      download(out.offerDocx, `Оферта (с изменениями) ${stamp}.docx`);
      download(out.combinedDocx, `Перечень изменений ${stamp}.docx`);
      setApplyReport(out.results);
      setStage("done");
    } catch (e) {
      setError(e instanceof Error ? e.message : "Не удалось собрать документы");
    } finally {
      setBusy(null);
    }
  };

  const stats = result?.stats;
  const activeCount = useMemo(
    () => items.filter((i) => !disabled.has(i.op.id) && i.op.type !== "manual").length,
    [items, disabled],
  );

  return (
    <div className="page">
      <header className="hero">
        <h1>genOferta AI</h1>
        <p>
          Объединение изменений в Публичную оферту. Текст документа правит детерминированный
          движок — ИИ только разбирает формулировки инструкций и подсказывает координаты правки
          там, где формальные правила не справились.
        </p>
      </header>

      <div className="steps">
        <span className={stage === "upload" ? "on" : ""}>1. Документы</span>
        <span className={stage === "review" ? "on" : ""}>2. Проверка правок</span>
        <span className={stage === "done" ? "on" : ""}>3. Готово</span>
      </div>

      {error && <div className="notice err">{error}</div>}

      {stage === "upload" && (
        <>
          <div className="card">
            <h2>Оферта — текущая редакция</h2>
            <p className="hint">
              Файл, который нужно изменить (например, «Версия 70.docx»). Один документ.
            </p>
            <FileDrop label="Оферта (.docx)" files={offer} onChange={setOffer} />
          </div>

          <div className="card">
            <h2>Документы «Изменения»</h2>
            <p className="hint">
              Нормативные документы с перечнем правок. Можно загрузить сразу несколько — они
              будут применены в порядке следования пунктов Оферты, а не в порядке загрузки.
            </p>
            <FileDrop
              multiple
              label="Изменения (.docx)"
              files={changes}
              onChange={setChanges}
            />
          </div>

          <div className="card">
            <h2>Режим разбора</h2>
            <p className="hint">
              {health === null
                ? "Проверяем доступность ИИ…"
                : health.configured
                  ? `Ключ настроен. Модель: ${health.model}`
                  : "Работает только алгоритмический разбор — см. ниже."}
            </p>

            {health && !health.configured && health.message && (
              <div className="notice err" style={{ marginBottom: 12 }}>
                {health.message}
              </div>
            )}

            {health && (
              <div className="row" style={{ marginBottom: 12 }}>
                <button
                  className="tiny"
                  disabled={probing}
                  onClick={async () => {
                    setProbing(true);
                    setProbe(null);
                    try {
                      const r = await fetch("/api/ai/health?probe=1");
                      const d = (await r.json()) as AiHealth & { probe?: KeyProbe };
                      setHealth(d);
                      setProbe(d.probe ?? { ok: false, message: "Проверка не выполнена" });
                    } catch (e) {
                      setProbe({
                        ok: false,
                        message: e instanceof Error ? e.message : "Проверка не выполнена",
                      });
                    } finally {
                      setProbing(false);
                    }
                  }}
                >
                  {probing ? "Проверяем…" : "Проверить ключ у OpenRouter"}
                </button>
                <span style={{ color: "var(--muted)", fontSize: 12 }}>
                  Мгновенно и без расхода токенов — спрашиваем OpenRouter, принимает ли он ключ.
                </span>
              </div>
            )}

            {probe && (
              <div className={`notice ${probe.ok ? "info" : "err"}`}>
                {probe.ok ? "✅ " : "❌ "}
                {probe.message}
                {probe.label ? ` (ключ «${probe.label}»)` : ""}
              </div>
            )}
            <div className="grid2">
              <label className="toggle">
                <input
                  type="checkbox"
                  checked={useAi}
                  disabled={!health?.configured}
                  onChange={(e) => setUseAi(e.target.checked)}
                />
                <span>
                  <b>Подключить ИИ к разбору</b>
                  <small>
                    Модель зовут точечно — туда, где формальные правила не разобрали
                    формулировку или не уверены в ней.
                  </small>
                </span>
              </label>
              <label className="toggle">
                <input
                  type="checkbox"
                  checked={verifyConfident}
                  disabled={!useAi || !health?.configured}
                  onChange={(e) => setVerifyConfident(e.target.checked)}
                />
                <span>
                  <b>Второе мнение по каждой правке</b>
                  <small>
                    Модель проверяет и те правки, что алгоритм разобрал уверенно. Медленнее,
                    зато ловит случай «правило сработало, но не то».
                  </small>
                </span>
              </label>
              <label className="toggle">
                <input
                  type="checkbox"
                  checked={showOld}
                  onChange={(e) => setShowOld(e.target.checked)}
                />
                <span>
                  <b>Показывать прежний текст зачёркнутым</b>
                  <small>Как в образце: старая редакция остаётся рядом с новой.</small>
                </span>
              </label>
              <label className="toggle">
                <input
                  type="checkbox"
                  checked={highlightMode === "tracked"}
                  onChange={(e) => setHighlightMode(e.target.checked ? "tracked" : "color")}
                />
                <span>
                  <b>Режим рецензирования Word</b>
                  <small>
                    Правки как «Исправления» Word вместо выделения цветом — их можно принять
                    или отклонить средствами Word.
                  </small>
                </span>
              </label>
            </div>
          </div>

          <div className="row">
            <button
              className="primary"
              disabled={!offer.length || !changes.length || Boolean(busy)}
              onClick={() => void analyze()}
            >
              {busy ?? "Разобрать изменения"}
            </button>
            <span style={{ color: "var(--muted)", fontSize: 13 }}>
              Документы не покидают браузер. На сервер уходит только текст инструкции и карта
              структуры Оферты — и только если ИИ включён.
            </span>
          </div>

          {busy && progress.total > 0 && (
            <div className="progress">
              <i style={{ width: `${(progress.done / progress.total) * 100}%` }} />
            </div>
          )}
        </>
      )}

      {stage === "review" && result && (
        <>
          <div className="card">
            <h2>Найдено правок: {items.length}</h2>
            <p className="hint">
              Включено в сборку: {activeCount}. Проверьте расхождения и правки, помеченные
              «не применяется», — именно в них прячутся ошибки разбора.
            </p>
            <div className="stats">
              {(Object.keys(PROVENANCE) as Provenance[])
                .filter((k) => (stats?.[k] ?? 0) > 0)
                .map((k) => (
                  <span key={k} className={`badge ${PROVENANCE[k].cls}`} title={PROVENANCE[k].title}>
                    {PROVENANCE[k].label}: {stats?.[k]}
                  </span>
                ))}
            </div>
            {result.aiErrors.length > 0 && (
              <div className="notice warn">
                ИИ отвечал с ошибками, часть правок разобрана только алгоритмом:
                <br />
                {result.aiErrors.join("; ")}
              </div>
            )}
            {!result.aiUsed && useAi && (
              <div className="notice info">
                ИИ не понадобился — все формулировки разобраны формальными правилами.
              </div>
            )}
          </div>

          {items.map((item) => (
            <OperationCard
              key={item.op.id}
              item={item}
              enabled={!disabled.has(item.op.id)}
              onToggle={(on) =>
                setDisabled((prev) => {
                  const next = new Set(prev);
                  if (on) next.delete(item.op.id);
                  else next.add(item.op.id);
                  return next;
                })
              }
              onEdit={(op: Operation) =>
                setItems((prev) => prev.map((i) => (i.op.id === op.id ? { ...i, op } : i)))
              }
              onPickAlternative={() =>
                setItems((prev) =>
                  prev.map((i) =>
                    i.op.id === item.op.id && i.alternative
                      ? { ...i, op: { ...i.alternative, id: i.op.id }, alternative: i.op }
                      : i,
                  ),
                )
              }
            />
          ))}

          <div className="row" style={{ marginTop: 18 }}>
            <button className="ghost" onClick={() => setStage("upload")}>
              ← Назад
            </button>
            <div className="spacer" />
            <button className="primary" disabled={!activeCount || Boolean(busy)} onClick={() => void build()}>
              {busy ?? `Собрать документы (${activeCount})`}
            </button>
          </div>
        </>
      )}

      {stage === "done" && (
        <>
          <div className="card">
            <h2>Готово</h2>
            <p className="hint">
              Скачаны два файла: Оферта с выделенными изменениями и перечень изменений в порядке
              следования пунктов. Правки, отмеченные «вручную», в файлы не вошли — внесите их в
              Word самостоятельно.
            </p>
            {applyReport && (
              <ul className="filelist">
                {applyReport.map((r) => (
                  <li key={r.operationId}>
                    <span>{r.ok ? "✅" : "⚠️"}</span>
                    <span className="name">{r.message}</span>
                  </li>
                ))}
              </ul>
            )}
          </div>
          <div className="row">
            <button onClick={() => setStage("review")}>← К списку правок</button>
            <button
              className="ghost"
              onClick={() => {
                setStage("upload");
                setChanges([]);
                setResult(null);
                setItems([]);
                setApplyReport(null);
              }}
            >
              Начать заново
            </button>
          </div>
        </>
      )}
    </div>
  );
}
