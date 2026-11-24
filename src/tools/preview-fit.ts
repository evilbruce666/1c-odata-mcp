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
 *   1 — машинная часть целиком, сокращены читаемые таблицы view (linesComplete=false, linesShown), если так
 *       помещается: клиент, который читает строки из payload, получает все;
 *   2 — строки машинной части, повторённые в view.tables, не повторяются (rowsInView); в summary — итоги по всем;
 *   3 — показаны первые строки таблиц, место делится поровну (linesComplete=false, linesShown).
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

/**
 * Заметки строк с одинаковым текстом — одной заметкой: «Строка 1: X», …, «Строка 50: X» → «Строки 1–50: X». Оплата
 * на 50 договоров давала 100 заметок (две на строку) — 11 800 знаков, и при пределе 10 000 ответ не помещался.
 * Порядок — по первому появлению; прочие заметки — как есть.
 */
export function compactLineNotes(notes: readonly string[]): string[] {
  const groups = new Map<string, number[]>();
  const order: Array<string | { text: string }> = [];
  for (const n of notes) {
    const m = /^Строка (\d+): ([\s\S]*)$/.exec(n);
    if (!m) {
      order.push(n);
      continue;
    }
    const text = m[2]!;
    if (!groups.has(text)) {
      groups.set(text, []);
      order.push({ text });
    }
    groups.get(text)!.push(Number(m[1]));
  }
  const ranges = (lines: number[]) => {
    const sorted = [...new Set(lines)].sort((a, b) => a - b);
    const out: string[] = [];
    for (let i = 0; i < sorted.length; i++) {
      let j = i;
      while (j + 1 < sorted.length && sorted[j + 1] === sorted[j]! + 1) j++;
      out.push(j > i ? `${sorted[i]}–${sorted[j]}` : String(sorted[i]));
      i = j;
    }
    return out.join(", ");
  };
  return order.map((o) => {
    if (typeof o === "string") return o;
    const lines = groups.get(o.text)!;
    return `${new Set(lines).size > 1 ? "Строки" : "Строка"} ${ranges(lines)}: ${o.text}`;
  });
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
  if (size(sc) + added <= limit) return sc;
  // Не предпросмотр (ответ подтверждения) — сводятся только одинаковые заметки строк: строки документа в нём не
  // повторяются, а заметки на 50 строк оплаты — 11 800 знаков.
  if (sc["dryRun"] !== true) {
    if (!Array.isArray(sc["notes"])) return sc;
    const compact = compactLineNotes(sc["notes"] as string[]);
    return compact.length < (sc["notes"] as string[]).length ? { ...sc, notes: compact } : sc;
  }
  const rowsKey = ROW_KEYS.find((k) => sc[k] && typeof sc[k] === "object" && !Array.isArray(sc[k]));
  if (!rowsKey) return sc;
  const rows = sc[rowsKey] as Row;
  const tables = Object.keys(rows).filter((k) => Array.isArray(rows[k]) && (rows[k] as unknown[]).length > 0);
  if (!tables.length) return sc;
  const budget = fittedBudget(limit);
  const view = sc["view"] as PreviewView | undefined;
  // Ответ сокращается — одинаковые заметки строк сводятся в одну (ответ, который помещается, — как есть).
  const notes = Array.isArray(sc["notes"]) ? compactLineNotes(sc["notes"] as string[]) : [];
  const keyOf = new Map(tables.map((k) => [fieldLabel(k), k]));

  /**
   * Одна попытка. payloadFirst — машинная часть целиком, сокращаются только читаемые таблицы view: клиент, который
   * берёт строки из payload, получает все (выбор владельца 24.11). Иначе — строки, видные в view целиком, не
   * повторяются в машинной части (rowsInView), и сокращаются все списки.
   */
  const attempt = (payloadFirst: boolean): { out: Record<string, unknown>; fits: boolean } => {
    const viewTables: Record<string, Row[]> = { ...(view?.tables ?? {}) };
    const inView = payloadFirst
      ? []
      : tables.filter((k) => viewTables[fieldLabel(k)]?.length === (rows[k] as unknown[]).length);
    const machine: Row = Object.fromEntries(Object.entries(rows).filter(([k]) => !inView.includes(k)));
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
                payloadFirst ? `В ${rowsKey} — все строки, сокращены только view.tables. ` : ""
              }${
                rowsKey === "payload" ? "Подтверждение создаст документ" : "Подтверждение запишет таблицы"
              } со всеми строками из аргументов вызова.`,
            ]
          : []),
      ],
    });
    let out = build();
    if (!payloadFirst && size(out) <= budget) return { out, fits: true };

    // Показываемые списки: читаемые таблицы view и (кроме payloadFirst) оставшиеся машинные. Место делится поровну:
    // списки короче доли — целиком, остальные — до равной доли остатка. Раньше резался самый большой список до места,
    // оставшегося после всех прочих целиком, и при больших «Сериях» «Товарам» доставалась одна строка.
    const shown: Record<string, number> = {};
    const lists = (): Array<{ name: string; list: Row[]; set: (l: Row[]) => void }> => [
      ...Object.keys(viewTables).map((label) => ({
        name: keyOf.get(label) ?? label,
        list: viewTables[label]!,
        set: (l: Row[]) => (viewTables[label] = l),
      })),
      ...(payloadFirst ? [] : tables.filter((k) => k in machine)).map((k) => ({
        name: k,
        list: machine[k] as Row[],
        set: (l: Row[]) => (machine[k] = l),
      })),
    ];
    for (let guard = 0; guard < 20; guard++) {
      out = build({ linesComplete: false, linesShown: { ...shown } });
      const total = size(out);
      if (total <= budget) break;
      // Список из одной строки не сокращается: в делении он — постоянная часть, иначе на него уходила бы доля, а
      // остальные урезались бы по крупицам и не уложились (ревью: строки шире доли).
      const all = lists()
        .filter((l) => l.list.length > 1)
        .map((l) => ({ ...l, chars: size(l.list) }))
        .sort((a, b) => a.chars - b.chars);
      let room = budget - (total - all.reduce((t, l) => t + l.chars, 0));
      let cut = false;
      for (let i = 0; i < all.length; i++) {
        const l = all[i]!;
        const share = Math.max(room, 0) / (all.length - i);
        if (l.chars <= share) {
          room -= l.chars;
          continue;
        }
        const kept = fitToChars(l.list, 0, share);
        room -= size(kept);
        if (kept.length >= l.list.length) continue;
        l.set(kept);
        shown[l.name] = Math.min(shown[l.name] ?? kept.length, kept.length);
        cut = true;
      }
      if (!cut) break;
    }
    if (Object.keys(shown).length) out = build({ linesComplete: false, linesShown: shown });
    return { out, fits: size(out) <= budget };
  };

  // Сначала — машинная часть целиком, если она помещается (сокращается только view); иначе — прежний путь.
  const first = view?.tables && Object.keys(view.tables).length ? attempt(true) : undefined;
  return first?.fits ? first.out : attempt(false).out;
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
