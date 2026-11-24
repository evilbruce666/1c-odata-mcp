import type { CallToolResult } from "@modelcontextprotocol/server";
import { fitToChars, MAX_RESPONSE_CHARS, renderJson } from "./_shared.js";
import { fieldLabel, type PreviewView } from "./preview-view.js";

/**
 * Предпросмотр записи документа в пределе ответа (ODATA_MAX_RESPONSE_CHARS). Строки табличных частей в нём дважды —
 * машинные (payload, у изменения fields) и читаемые (view.tables), и документ на сотню строк его превышал: клиент
 * обрезал ответ и терял operationId, без которого подтверждение не пройдёт. Подтверждение собирает документ заново из
 * аргументов вызова, поэтому сокращается только показ, не запись.
 *
 * Уровни, пока ответ не уложится:
 *   0 — как есть, если помещается (вместе с operationId, который обёртка добавит);
 *   1 — строки машинной части, повторённые в view.tables, не повторяются; в summary — итоги по всем строкам;
 *   2 — показаны первые строки самых больших таблиц (linesComplete=false, linesShown).
 */

/** Поле operationId, которое обёртка записи добавит к ответу: `"operationId":"<uuid>",` — ровно 53 знака. */
export const OPERATION_ID_CHARS = 53;

/** Запас к пределу сокращённого ответа: operationId, примечание о сокращении. */
export const RESPONSE_RESERVE = 1_500;

/** Бюджет сокращённого ответа: как у страниц чтения (запас под служебные поля) минус запас под примечания. */
export const fittedBudget = (limit: number): number =>
  (limit >= MAX_RESPONSE_CHARS ? limit : limit - 1_000) - RESPONSE_RESERVE;

/** Итоги строк: суммы — в копейках, количество — в тысячных, без накопления ошибки округления. */
const TOTALS: Array<[string, number]> = [
  ["Количество", 1_000],
  ["Сумма", 100],
  ["СуммаНДС", 100],
  ["СуммаСНДС", 100],
];

type Row = Record<string, unknown>;

/**
 * Итоги таблицы: lines — число строк, поле итога — сумма по строкам, где оно числом. Если поле есть не во всех строках,
 * withValue называет, по скольким строкам сложено, — иначе неполный итог выглядел бы полным.
 */
export function linesSummary(
  rows: Row,
  tables: readonly string[],
): Record<string, Record<string, number | Record<string, number>>> {
  return Object.fromEntries(
    tables.map((k) => {
      const list = rows[k] as Row[];
      const out: Record<string, number | Record<string, number>> = { lines: list.length };
      const withValue: Record<string, number> = {};
      for (const [field, scale] of TOTALS) {
        const values = list.map((r) => r[field]).filter((v): v is number => typeof v === "number");
        if (!values.length) continue;
        out[field] = values.reduce((t, v) => t + Math.round(v * scale), 0) / scale;
        if (values.length < list.length) withValue[field] = values.length;
      }
      if (Object.keys(withValue).length) out["withValue"] = withValue;
      return [k, out];
    }),
  );
}

/** Машинная часть предпросмотра: payload у создания, fields у изменения. */
const ROW_KEYS = ["payload", "fields"] as const;

export function fitDocumentPreview(
  sc: Record<string, unknown>,
  limit: number,
  /** Символов, которые обёртка добавит к ответу (operationId); 0 — не добавит. */
  added: number,
): Record<string, unknown> {
  const size = (o: unknown) => renderJson(o).length;
  if (sc["dryRun"] !== true || size(sc) + added <= limit) return sc;
  const rowsKey = ROW_KEYS.find((k) => sc[k] && typeof sc[k] === "object" && !Array.isArray(sc[k]));
  if (!rowsKey) return sc;
  const rows = sc[rowsKey] as Row;
  const tables = Object.keys(rows).filter((k) => Array.isArray(rows[k]) && (rows[k] as unknown[]).length > 0);
  if (!tables.length) return sc;
  const budget = fittedBudget(limit);
  const view = sc["view"] as PreviewView | undefined;
  const viewTables: Record<string, Row[]> = { ...(view?.tables ?? {}) };
  // Таблица видна в view целиком — машинную копию не повторяем.
  const inView = tables.filter((k) => viewTables[fieldLabel(k)]?.length === (rows[k] as unknown[]).length);
  const machine: Row = Object.fromEntries(Object.entries(rows).filter(([k]) => !inView.includes(k)));
  const notes = Array.isArray(sc["notes"]) ? [...(sc["notes"] as string[])] : [];
  const build = (extra: Row = {}) => ({
    ...sc,
    ...(view ? { view: { ...view, tables: viewTables } } : {}),
    [rowsKey]: machine,
    summary: linesSummary(rows, tables),
    ...(inView.length ? { rowsInView: inView } : {}),
    ...extra,
    notes: [
      ...notes,
      ...(inView.length
        ? [`Строки ${inView.join(", ")} не повторены в ${rowsKey}: они в view.tables, итоги — в summary.`]
        : []),
      ...(extra["linesComplete"] === false
        ? [
            `Документ большой: показаны первые строки (linesShown), итоги summary — по всем строкам. ${
              rowsKey === "payload" ? "Подтверждение создаст документ" : "Подтверждение запишет таблицы"
            } со всеми строками из аргументов вызова.`,
          ]
        : []),
    ],
  });
  let out = build();
  if (size(out) <= budget) return out;

  // Показываемые списки: читаемые таблицы view и оставшиеся машинные. Режем самый большой, пока не уложится.
  const keyOf = new Map(tables.map((k) => [fieldLabel(k), k]));
  const shown: Record<string, number> = {};
  const lists = (): Array<{ name: string; list: Row[]; set: (l: Row[]) => void }> => [
    ...Object.keys(viewTables).map((label) => ({
      name: keyOf.get(label) ?? label,
      list: viewTables[label]!,
      set: (l: Row[]) => (viewTables[label] = l),
    })),
    ...tables
      .filter((k) => k in machine)
      .map((k) => ({ name: k, list: machine[k] as Row[], set: (l: Row[]) => (machine[k] = l) })),
  ];
  for (let guard = 0; guard < 20; guard++) {
    out = build({ linesComplete: false, linesShown: { ...shown } });
    const total = size(out);
    if (total <= budget) break;
    const biggest = lists().sort((a, b) => size(b.list) - size(a.list))[0];
    if (!biggest || biggest.list.length <= 1) break;
    const room = budget - (total - size(biggest.list));
    const kept = fitToChars(biggest.list, 0, Math.max(room, 0));
    if (kept.length >= biggest.list.length) break;
    biggest.set(kept);
    shown[biggest.name] = Math.min(shown[biggest.name] ?? kept.length, kept.length);
  }
  return Object.keys(shown).length ? build({ linesComplete: false, linesShown: shown }) : out;
}

/** Обёртка сервера: подогнать предпросмотр записи и перерисовать текст ответа. */
export function fitWritePreview(result: CallToolResult, limit: number, added: number): CallToolResult {
  const sc = result.structuredContent as Record<string, unknown> | undefined;
  if (result.isError || !sc) return result;
  const fitted = fitDocumentPreview(sc, limit, added);
  if (fitted === sc) return result;
  return {
    ...result,
    structuredContent: fitted,
    content: result.content.map((c) => (c.type === "text" ? { ...c, text: renderJson(fitted) } : c)),
  };
}
