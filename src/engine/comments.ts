// Комментарии Word (область рецензирования) для правок, которые движок не
// вносит сам.
//
// Комментарий в OOXML состоит из двух частей, и обе обязательны:
//   • запись <w:comment> в отдельной части word/comments.xml;
//   • якорь в тексте — commentRangeStart / commentRangeEnd вокруг фрагмента и
//     run с <w:commentReference> сразу после него.
// Часть без якоря Word просто не покажет — именно так и выглядит наследие
// OnlyOffice в некоторых редакциях Оферты: часть с комментарием есть, а в
// document.xml нет ни одного commentRangeStart, поэтому в Word пусто.
//
// Часть подключается связью СТАНДАРТНОГО типа. Тот же OnlyOffice подключает
// свою часть типом http://schemas.onlyoffice.com/commentsDocument, который Word
// за комментарии не считает, поэтому опираться на существующую связь нельзя —
// свою часть заводим сами и связываем как положено.
import { escapeXml } from "./ooxml";

/** Стандартный тип связи части комментариев. */
const REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const PART_PATH = "word/comments.xml";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export interface CommentsState {
  /** Содержимое word/comments.xml (создаётся, если части не было). */
  xml: string;
  /** word/_rels/document.xml.rels с гарантированной связью на часть. */
  rels: string;
  /** [Content_Types].xml с override для части. */
  contentTypes: string;
  /** Следующий свободный номер комментария. */
  nextId: number;
  /** Были ли изменения — чтобы не переписывать части зря. */
  touched: boolean;
}

const EMPTY_COMMENTS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:comments xmlns:w="${W_NS}"></w:comments>`;

/**
 * Подготовить состояние комментариев: часть, связь и тип содержимого.
 * Ничего не пишет — только считает, каким всё должно стать.
 */
export function prepareComments(
  existingXml: string | null,
  rels: string,
  contentTypes: string,
): CommentsState {
  const xml = existingXml ?? EMPTY_COMMENTS;
  return {
    xml,
    rels: ensureRelationship(rels),
    contentTypes: ensureOverride(contentTypes),
    nextId: maxCommentId(xml) + 1,
    touched: false,
  };
}

/** Наибольший занятый номер комментария (−1, если комментариев нет). */
function maxCommentId(xml: string): number {
  let max = -1;
  const re = /<w:comment\b[^>]*\sw:id="(\d+)"/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) max = Math.max(max, parseInt(m[1], 10));
  return max;
}

/** Свободный Id связи вида rIdN. */
function freeRelId(rels: string): string {
  let max = 0;
  const re = /Id="rId(\d+)"/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(rels)) !== null) max = Math.max(max, parseInt(m[1], 10));
  return `rId${max + 1}`;
}

function ensureRelationship(rels: string): string {
  // Связь именно СТАНДАРТНОГО типа: чужая связь на ту же часть (OnlyOffice)
  // Word не устраивает, а две связи разных типов друг другу не мешают.
  if (new RegExp(`Type="${REL_TYPE}"`).test(rels)) return rels;
  const rel =
    `<Relationship Id="${freeRelId(rels)}" Type="${REL_TYPE}" Target="comments.xml"/>`;
  return rels.replace(/<\/Relationships>\s*$/, `${rel}</Relationships>`);
}

function ensureOverride(contentTypes: string): string {
  if (contentTypes.includes(`PartName="/${PART_PATH}"`)) return contentTypes;
  const ov = `<Override PartName="/${PART_PATH}" ContentType="${CONTENT_TYPE}"/>`;
  return contentTypes.replace(/<\/Types>\s*$/, `${ov}</Types>`);
}

/**
 * Добавить комментарий и вернуть его номер. Текст разбивается на абзацы по
 * переводам строки — так список данных к внесению читается в панели построчно.
 */
export function addComment(
  state: CommentsState,
  text: string,
  author = "Объединение изменений",
  initials = "ОИ",
): number {
  const id = state.nextId;
  const date = new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
  const paras = text
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .map(
      (line) =>
        `<w:p><w:r><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`,
    )
    .join("");
  const comment =
    `<w:comment w:id="${id}" w:author="${escapeXml(author)}" ` +
    `w:initials="${escapeXml(initials)}" w:date="${date}">` +
    (paras || "<w:p/>") +
    `</w:comment>`;
  state.xml = state.xml.replace(/<\/w:comments>\s*$/, `${comment}</w:comments>`);
  state.nextId = id + 1;
  state.touched = true;
  return id;
}

/** Открывающий якорь комментария (ставится перед комментируемым фрагментом). */
export function commentRangeStart(id: number): string {
  return `<w:commentRangeStart w:id="${id}"/>`;
}

/**
 * Закрывающий якорь вместе со ссылкой. Ссылка обязана лежать ВНУТРИ run’а:
 * голый <w:commentReference> вне run’а делает документ невалидным, и Word
 * показывает «в файле обнаружено содержимое, которое не удалось прочитать».
 */
export function commentRangeEnd(id: number): string {
  // Без rStyle="CommentReference": такого стиля в styles.xml Оферты может не
  // быть, а висячая ссылка на стиль — ровно то, из-за чего ранее не открывался
  // объединённый файл изменений. Формат значка Word подставит сам.
  return (
    `<w:commentRangeEnd w:id="${id}"/>` +
    `<w:r><w:commentReference w:id="${id}"/></w:r>`
  );
}
