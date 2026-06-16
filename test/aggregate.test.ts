import { describe, it, expect } from "vitest";
import type { Connection } from "../src/context.js";
import {
  collectDocuments,
  fetchAllForAggregation,
  AggregateOverflowError,
  splitByMonth,
} from "../src/odata/aggregate.js";

/**
 * Фейковый OData-клиент: отдаёт N строк постранично ($top/$skip) и уважает
 * фильтр по дате (Date ge/le) — этого достаточно, чтобы проверить, что
 * агрегатор суммирует ВСЁ (а не обрезает на 1000) и корректно режет по месяцам.
 */
function fakeConn(rows: Array<{ Date: string }>, cap: number): Connection {
  const client = {
    async getCollection(path: string): Promise<{ value: unknown[] }> {
      const dec = decodeURIComponent(path);
      const ge = /Date ge datetime'([^']+)'/.exec(dec)?.[1];
      const le = /Date le datetime'([^']+)'/.exec(dec)?.[1];
      const top = Number(/\$top=(\d+)/.exec(path)?.[1] ?? rows.length);
      const skip = Number(/\$skip=(\d+)/.exec(path)?.[1] ?? 0);
      let pool = rows;
      if (ge) pool = pool.filter((r) => r.Date >= ge);
      if (le) pool = pool.filter((r) => r.Date <= le);
      return { value: pool.slice(skip, skip + top) };
    },
  };
  return { client, behavior: { analyticsMaxRows: cap, pageSize: 100 } } as unknown as Connection;
}

const makeRows = (n: number, date = "2025-06-15T00:00:00"): Array<{ Date: string }> =>
  Array.from({ length: n }, () => ({ Date: date }));

describe("fetchAllForAggregation", () => {
  it("суммирует ВСЕ строки выше старого лимита 1000 (регресс бага усечки)", async () => {
    const conn = fakeConn(makeRows(1500), 200_000);
    const { rows, meta } = await fetchAllForAggregation(conn, "Doc", {});
    expect(rows.length).toBe(1500);
    expect(meta.rowsScanned).toBe(1500);
    expect(meta.chunks).toBe(1);
  });

  it("бросает громкую ошибку при переполнении потолка (не тихая обрезка)", async () => {
    const conn = fakeConn(makeRows(50), 10);
    await expect(fetchAllForAggregation(conn, "Doc", {})).rejects.toBeInstanceOf(AggregateOverflowError);
  });
});

describe("collectDocuments", () => {
  it("берёт весь период одним окном, если влезает", async () => {
    const conn = fakeConn(makeRows(1500), 200_000);
    const { rows, meta } = await collectDocuments(conn, "Doc", {
      dateField: "Date",
      from: "2025-01-01",
      to: "2025-12-31",
      select: ["Date"],
    });
    expect(rows.length).toBe(1500);
    expect(meta.chunks).toBe(1);
  });

  it("при переполнении режет по месяцам и суммирует ПОЛНОСТЬЮ", async () => {
    const rows = [
      ...makeRows(30, "2025-01-10T00:00:00"),
      ...makeRows(30, "2025-02-10T00:00:00"),
      ...makeRows(30, "2025-03-10T00:00:00"),
    ];
    const conn = fakeConn(rows, 50); // весь период (90) > 50, но каждый месяц (30) ≤ 50
    const res = await collectDocuments(conn, "Doc", {
      dateField: "Date",
      from: "2025-01-01",
      to: "2025-03-31",
      select: ["Date"],
    });
    expect(res.rows.length).toBe(90); // ничего не потеряли
    expect(res.meta.chunks).toBe(3); // три окна
  });

  it("бросает, если даже один месяц не влезает", async () => {
    const conn = fakeConn(makeRows(100, "2025-02-10T00:00:00"), 50);
    await expect(
      collectDocuments(conn, "Doc", {
        dateField: "Date",
        from: "2025-02-01",
        to: "2025-02-28",
        select: ["Date"],
      }),
    ).rejects.toBeInstanceOf(AggregateOverflowError);
  });
});

describe("splitByMonth", () => {
  it("режет период по календарным месяцам с клампом к границам", () => {
    expect(splitByMonth("2025-01-15", "2025-03-10")).toEqual([
      { from: "2025-01-15", to: "2025-01-31" },
      { from: "2025-02-01", to: "2025-02-28" },
      { from: "2025-03-01", to: "2025-03-10" },
    ]);
  });

  it("один месяц — одно окно", () => {
    expect(splitByMonth("2025-06-01", "2025-06-30")).toEqual([{ from: "2025-06-01", to: "2025-06-30" }]);
  });
});
