"use client";

import { useState } from "react";
import { OP_TYPE_LABEL, PROVENANCE, targetLabel } from "@/lib/labels";
import type { HybridOperation } from "@/ai/types";
import type { Operation } from "@/engine/types";

export default function OperationCard({
  item,
  enabled,
  onToggle,
  onEdit,
  onPickAlternative,
}: {
  item: HybridOperation;
  enabled: boolean;
  onToggle: (on: boolean) => void;
  onEdit: (op: Operation) => void;
  onPickAlternative: () => void;
}) {
  const [open, setOpen] = useState(
    // Правки, которые не применятся или требуют выбора, раскрыты сразу:
    // именно их оператор обязан посмотреть, а не те, что и так в порядке.
    item.provenance === "conflict" ||
      item.provenance === "manual" ||
      item.dryRun?.ok === false,
  );
  const [draft, setDraft] = useState(item.op.payload ?? "");
  const p = PROVENANCE[item.provenance];
  const failed = item.dryRun && !item.dryRun.ok;

  return (
    <div className={`op${enabled ? "" : " off"}${failed ? " bad" : ""}`}>
      <div className="op-head">
        <input
          type="checkbox"
          checked={enabled}
          disabled={item.op.type === "manual"}
          onChange={(e) => onToggle(e.target.checked)}
          title={
            item.op.type === "manual"
              ? "Эту правку движок не вносит — её нужно внести в Word вручную"
              : "Включить правку в сборку"
          }
        />
        <div className="op-title">
          <div className="line1">
            <span className="target">{targetLabel(item.op)}</span>
            <span className={`badge ${p.cls}`} title={p.title}>
              {p.label}
            </span>
            {failed && (
              <span className="badge manual" title={item.dryRun?.message}>
                не применяется
              </span>
            )}
          </div>
          <div className="type">
            {OP_TYPE_LABEL[item.op.type]} · {item.op.sourceDoc}
          </div>
        </div>
        <button className="tiny ghost" onClick={() => setOpen((v) => !v)}>
          {open ? "свернуть" : "подробно"}
        </button>
      </div>

      {open && (
        <div className="op-body">
          <div>
            <div className="field-label">Текст инструкции (как в документе «Изменения»)</div>
            <div className="quote src">{item.op.rawText}</div>
          </div>

          {item.aiReasoning && (
            <div>
              <div className="field-label">Как это прочитал ИИ</div>
              <div className="quote ai">{item.aiReasoning}</div>
            </div>
          )}

          {failed && (
            <div>
              <div className="field-label">Почему правка не применяется</div>
              <div className="quote err">{item.dryRun?.message}</div>
            </div>
          )}

          {item.op.note && (
            <div>
              <div className="field-label">Примечание</div>
              <div className="quote warn">{item.op.note}</div>
            </div>
          )}

          {item.op.warnings?.length ? (
            <div>
              <div className="field-label">Предупреждения</div>
              <div className="quote warn">{item.op.warnings.join("\n")}</div>
            </div>
          ) : null}

          {item.op.anchor && (
            <div>
              <div className="field-label">Якорь (слова, после/перед которыми вставка)</div>
              <div className="quote">{item.op.anchor}</div>
            </div>
          )}

          {item.op.find && (
            <div>
              <div className="field-label">Что ищем в тексте</div>
              <div className="quote">{item.op.find}</div>
            </div>
          )}

          {item.op.rows?.length ? (
            <div>
              <div className="field-label">Строки таблицы</div>
              <table className="rows">
                <tbody>
                  {item.op.rows.map((r, i) => (
                    <tr key={i}>
                      {r.map((c, j) => (
                        <td key={j}>{c}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : null}

          {item.op.payload !== undefined && (
            <div>
              <div className="field-label">
                Текст правки — правится вручную, если разбор ошибся в границах
              </div>
              <textarea
                value={draft}
                onChange={(e) => setDraft(e.target.value)}
                onBlur={() => {
                  if (draft !== item.op.payload) onEdit({ ...item.op, payload: draft });
                }}
              />
            </div>
          )}

          {item.alternative && (
            <div>
              <div className="field-label">
                Второй вариант прочтения ({item.provenance === "conflict" ? "от другого источника" : "альтернатива"})
              </div>
              <div className="quote ai">
                <b>{targetLabel(item.alternative)}</b> · {OP_TYPE_LABEL[item.alternative.type]}
                {item.alternative.payload ? `\n\n${item.alternative.payload}` : ""}
              </div>
              <button className="tiny" style={{ marginTop: 8 }} onClick={onPickAlternative}>
                Использовать этот вариант
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
