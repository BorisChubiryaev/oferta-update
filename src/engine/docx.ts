// Работа с .docx как с zip-контейнером OOXML.
import JSZip from "jszip";

export interface DocxParts {
  zip: JSZip;
  document: string; // word/document.xml
  footnotes: string | null; // word/footnotes.xml (может отсутствовать)
  numbering: string | null; // word/numbering.xml (для восстановления нумерации)
  styles: string | null; // word/styles.xml (нумерация может задаваться стилем)
  comments: string | null; // word/comments.xml (может отсутствовать)
  rels: string; // word/_rels/document.xml.rels
  contentTypes: string; // [Content_Types].xml
}

export async function loadDocx(data: Uint8Array | ArrayBuffer): Promise<DocxParts> {
  const zip = await JSZip.loadAsync(data);
  const read = async (path: string) => {
    const f = zip.file(path);
    return f ? await f.async("string") : null;
  };
  const document = (await read("word/document.xml"))!;
  return {
    zip,
    document,
    footnotes: await read("word/footnotes.xml"),
    numbering: await read("word/numbering.xml"),
    styles: await read("word/styles.xml"),
    comments: await read("word/comments.xml"),
    // Связи и типы содержимого нужны, чтобы подключить часть комментариев:
    // без них новая часть в пакете есть, но Word о ней не знает.
    rels: (await read("word/_rels/document.xml.rels")) ?? "",
    contentTypes: (await read("[Content_Types].xml")) ?? "",
  };
}

/** Записать изменённые части обратно и отдать байты .docx. */
export async function saveDocx(
  parts: DocxParts,
  patched: {
    document?: string;
    footnotes?: string;
    comments?: string;
    rels?: string;
    contentTypes?: string;
  },
): Promise<Uint8Array> {
  if (patched.document !== undefined) {
    parts.zip.file("word/document.xml", patched.document);
  }
  if (patched.footnotes !== undefined && patched.footnotes !== null) {
    parts.zip.file("word/footnotes.xml", patched.footnotes);
  }
  if (patched.comments) {
    parts.zip.file("word/comments.xml", patched.comments);
    if (patched.rels) parts.zip.file("word/_rels/document.xml.rels", patched.rels);
    if (patched.contentTypes) parts.zip.file("[Content_Types].xml", patched.contentTypes);
  }
  return parts.zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}
