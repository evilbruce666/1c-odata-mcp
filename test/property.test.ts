import fc from "fast-check";
import { describe, expect, it } from "vitest";
import type { Connection } from "../src/context.js";
import { buildQuery, contains, odataGuid, odataString } from "../src/odata/query.js";
import { readRegisterRows } from "../src/odata/register-rows.js";
import { netAccruals } from "../src/odata/tax-calendar.js";
import { oneLine } from "../src/tools/preview-view.js";

/**
 * Свойства на случайных входах (fast-check): там, где ошибка дорога — текст из 1С и от модели попадает в запрос к
 * 1С или на экран подтверждения, а суммы налогов не должны теряться. Каждый тест прогоняет сотни входов, включая
 * кавычки, суррогатные пары, управляющие и невидимые символы; при падении fast-check печатает минимальный пример.
 */

/**
 * Опасные символы — половина входа: на случайном Юникоде кавычка или символ направления текста почти не выпадают, и
 * поломка экранирования проходила бы незамеченной (проверено: без этой смеси тесты не ловили снятое экранирование).
 */
const special = fc.constantFrom(
  "'",
  '"',
  "%",
  "_",
  "[",
  "]",
  "\\",
  "&",
  "=",
  "+",
  "#",
  "?",
  " ",
  "  ",
  "\n",
  "\r",
  "\t",
  "\u0000",
  "\u0085",
  "\u200B",
  "\u200E",
  "\u202E",
  "\u2066",
  "\u2028",
  "\uFEFF",
  "\u{1F600}",
  "Ё",
);
const unit = fc.oneof(special, fc.string({ unit: "binary", minLength: 1, maxLength: 1 }));
const anyText = fc.string({ unit, maxLength: 40 });

/** Строковый литерал OData обратно в значение: '' внутри — одна кавычка. */
function parseLiteral(literal: string): string {
  expect(literal.startsWith("'") && literal.endsWith("'") && literal.length >= 2).toBe(true);
  const inner = literal.slice(1, -1);
  // Одиночная кавычка внутри закрыла бы литерал — и всё после неё стало бы частью запроса.
  expect(inner.replace(/''/g, "")).not.toContain("'");
  return inner.replace(/''/g, "'");
}

describe("литералы запроса к 1С", () => {
  it("строка не выходит из кавычек и читается обратно без потерь", () => {
    fc.assert(
      fc.property(anyText, (s) => {
        expect(parseLiteral(odataString(s))).toBe(s);
      }),
    );
  });

  it("подстрока: символы ПОДОБНО экранированы, обратная косая в паре, литерал цел", () => {
    fc.assert(
      fc.property(anyText, (s) => {
        const m = /^substringof\((.*), Поле\)$/s.exec(contains("Поле", s));
        expect(m).not.toBeNull();
        const value = parseLiteral(m![1]!);
        // Шаблон ПОДОБНО обратно в текст: [x] — буквально x, \\ — одна косая.
        expect(value.replace(/\[([%_[])\]/g, "$1").replace(/\\\\/g, "\\")).toBe(s);
        expect(value.replace(/\[[%_[]\]/g, "")).not.toMatch(/[%_[]/);
        expect(value.replace(/\\\\/g, "")).not.toContain("\\");
      }),
    );
  });

  it("GUID: принимается только формат GUID, на выходе — без скобок и кавычек", () => {
    fc.assert(
      fc.property(fc.uuid(), fc.boolean(), fc.boolean(), (u, braces, upper) => {
        const raw = upper ? u.toUpperCase() : u;
        expect(odataGuid(braces ? `{${raw}}` : raw)).toBe(`guid'${raw}'`);
      }),
    );
    fc.assert(
      fc.property(anyText, (s) => {
        let out: string | undefined;
        try {
          out = odataGuid(s);
        } catch {
          return;
        }
        expect(out).toMatch(/^guid'[0-9a-fA-F-]{36}'$/);
      }),
    );
  });

  it("параметры запроса: любое значение отбора — ровно один $filter, читается обратно без потерь", () => {
    fc.assert(
      fc.property(anyText.filter(Boolean), fc.array(anyText, { maxLength: 3 }), (filter, select) => {
        const q = buildQuery({ filter, select: select.length ? select : undefined, top: 5 });
        expect(q.startsWith("?")).toBe(true);
        const params = q
          .slice(1)
          .split("&")
          .map((p) => p.split("="));
        const keys = params.map(([k]) => k);
        expect(keys).toEqual(["$format", ...(select.length ? ["$select"] : []), "$filter", "$top"]);
        for (const p of params) expect(p).toHaveLength(2);
        expect(decodeURIComponent(params.find(([k]) => k === "$filter")![1]!)).toBe(filter);
      }),
    );
  });
});

describe("read_register: значение отбора от модели", () => {
  const register = "InformationRegister_Проверка";
  const read = async (value: string) => {
    const paths: string[] = [];
    const conn = {
      behavior: { maxRows: 10 },
      available: async () => new Set([register]),
      getMetadata: async () => ({
        entities: new Map([
          [
            register,
            {
              entitySet: register,
              properties: [
                { name: "Описание", type: "Edm.String", nullable: true },
                { name: "Сумма", type: "Edm.Double", nullable: true },
              ],
            },
          ],
        ]),
      }),
      client: {
        getCollection: async (path: string) => {
          paths.push(path);
          return { value: [] };
        },
      },
    } as unknown as Connection;
    await readRegisterRows(conn, {
      register,
      select: ["Сумма"],
      filters: [
        { field: "Описание", op: "eq", value },
        { field: "Описание", op: "contains", value },
      ],
    });
    return paths;
  };

  it("остаётся литералом: один запрос, тот же набор параметров, отбор — ровно два условия", async () => {
    await fc.assert(
      fc.asyncProperty(anyText, async (value) => {
        const paths = await read(value);
        expect(paths).toHaveLength(1);
        const [set, query] = paths[0]!.split("?");
        expect(set).toBe(register);
        const params = query!.split("&").map((p) => p.split("="));
        expect(params.map(([k]) => k)).toEqual(["$format", "$select", "$filter", "$top", "$skip"]);
        expect(decodeURIComponent(params[2]![1]!)).toBe(
          `Описание eq ${odataString(value)} and ${contains("Описание", value)}`,
        );
      }),
    );
  });
});

describe("текст из 1С на экране подтверждения", () => {
  /* eslint-disable no-control-regex -- управляющие символы и есть то, чего в выводе быть не должно */
  const hidden = new RegExp(
    "[\\u0000-\\u001F\\u007F-\\u009F\\u200B-\\u200F\\u2028\\u2029\\u202A-\\u202E\\u2060-\\u2069\\uFEFF]",
  );
  /* eslint-enable no-control-regex */

  it("одна строка без управляющих, невидимых и символов направления; повтор ничего не меняет", () => {
    fc.assert(
      fc.property(fc.string({ unit, maxLength: 80 }), (s) => {
        const out = oneLine(s);
        expect(out).not.toMatch(hidden);
        expect(out).not.toMatch(/ {2}/);
        expect(out).toBe(out.trim());
        expect(oneLine(out)).toBe(out);
      }),
    );
  });

  it("обычный текст не искажается", () => {
    const plain = fc.stringMatching(/^[А-Яа-яЁёA-Za-z0-9.,«»"№-]+( [А-Яа-яЁёA-Za-z0-9.,«»"№-]+)*$/);
    fc.assert(
      fc.property(plain, (s) => {
        expect(oneLine(s)).toBe(s);
      }),
    );
  });
});

describe("свод начислений налогов", () => {
  const ref = (i: number) => `00000000-0000-4000-8000-${String(i).padStart(12, "0")}`;
  const EMPTY = "00000000-0000-0000-0000-000000000000";
  const row = fc.record({
    СчетУчета_Key: fc.constantFrom(ref(1), ref(2), ref(3)),
    // Сторно патента находит своё начисление на соседней дате — пару нужно часто порождать, иначе ветка не проверяется.
    ДокументПатент_Key: fc.constantFrom(EMPTY, ref(9)),
    СрокУплаты: fc.constantFrom("2026-12-28", "2027-01-11", "2027-03-29", "2027-04-27", "2027-04-28"),
    ПериодРасчета: fc.constantFrom("2026-12-31", "2027-01-31", "2027-03-31"),
    Начислено: fc.integer({ min: -50_000, max: 50_000 }).map((c) => c / 100),
    Active: fc.boolean(),
  });
  const cents = (n: number) => Math.round(n * 100);

  it("по каждому счёту итог свода равен сумме активных строк: сторно переносит, но не теряет", () => {
    fc.assert(
      fc.property(fc.array(row, { maxLength: 12 }), (rows) => {
        const raw = rows.map((r) => ({
          ...r,
          СрокУплаты: `${r.СрокУплаты}T00:00:00`,
          ПериодРасчета: `${r.ПериодРасчета}T00:00:00`,
        }));
        const { due, excess } = netAccruals(raw);
        for (const account of [ref(1), ref(2), ref(3)]) {
          const expected = rows
            .filter((r) => r.Active && r.СчетУчета_Key === account)
            .reduce((s, r) => s + cents(r.Начислено), 0);
          const got = [...due, ...excess]
            .filter((a) => a.accountKey === account)
            .reduce((s, a) => s + cents(a.amount), 0);
          expect(got).toBe(expected);
        }
        // Платёж — только положительный, отрицательный остаток — только в excess.
        expect(due.every((a) => a.amount > 0)).toBe(true);
        expect(excess.every((a) => a.amount < 0)).toBe(true);
      }),
      { numRuns: 2000 },
    );
  });
});
