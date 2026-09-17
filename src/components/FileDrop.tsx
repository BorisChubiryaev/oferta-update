"use client";

import { useRef, useState } from "react";

export interface PickedFile {
  name: string;
  data: Uint8Array;
}

/** Чтение .docx в память: документы никуда не отправляются, разбор — в браузере. */
async function readDocx(file: File): Promise<PickedFile> {
  return { name: file.name, data: new Uint8Array(await file.arrayBuffer()) };
}

export default function FileDrop({
  multiple,
  label,
  files,
  onChange,
}: {
  multiple?: boolean;
  label: string;
  files: PickedFile[];
  onChange: (files: PickedFile[]) => void;
}) {
  const input = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const accept = async (list: FileList | null) => {
    if (!list?.length) return;
    const picked = Array.from(list);
    const bad = picked.filter((f) => !f.name.toLowerCase().endsWith(".docx"));
    if (bad.length) {
      // .doc (старый формат) — не zip, и загрузчик docx на нём падает с
      // невнятной ошибкой. Сказать об этом здесь понятнее, чем на разборе.
      setError(
        `Нужен формат .docx. Не подходит: ${bad.map((f) => f.name).join(", ")}. ` +
          "Откройте файл в Word и сохраните как .docx.",
      );
      return;
    }
    setError(null);
    const read = await Promise.all(picked.map(readDocx));
    onChange(multiple ? [...files, ...read] : read.slice(0, 1));
  };

  return (
    <div>
      <div
        className={`drop${over ? " over" : ""}`}
        onClick={() => input.current?.click()}
        onDragOver={(e) => {
          e.preventDefault();
          setOver(true);
        }}
        onDragLeave={() => setOver(false)}
        onDrop={(e) => {
          e.preventDefault();
          setOver(false);
          void accept(e.dataTransfer.files);
        }}
      >
        <input
          ref={input}
          type="file"
          accept=".docx"
          multiple={multiple}
          onChange={(e) => {
            void accept(e.target.files);
            // Сброс нужен, чтобы повторный выбор того же файла снова дал событие.
            e.target.value = "";
          }}
        />
        <div>{label}</div>
        <div style={{ color: "var(--muted)", fontSize: 13, marginTop: 4 }}>
          перетащите сюда или нажмите для выбора
        </div>
      </div>

      {error && (
        <div className="notice err" style={{ marginTop: 10 }}>
          {error}
        </div>
      )}

      {files.length > 0 && (
        <ul className="filelist">
          {files.map((f, i) => (
            <li key={`${f.name}-${i}`}>
              <span>📄</span>
              <span className="name">{f.name}</span>
              <span style={{ color: "var(--muted)" }}>{Math.round(f.data.length / 1024)} КБ</span>
              <button
                className="tiny ghost"
                onClick={(e) => {
                  e.stopPropagation();
                  onChange(files.filter((_, k) => k !== i));
                }}
              >
                убрать
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
