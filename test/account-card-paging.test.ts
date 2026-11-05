import { afterEach, describe, expect, it, vi } from "vitest";
import { fetchMovements, STABLE_MOVEMENT_ORDER } from "../src/odata/account-card.js";
import { fake1C } from "./helpers/fake-1c.js";

const ACC = "aaaaaaaa-0000-4000-8000-000000000001";
const row = (doc: string, line = 1) => ({
  Period: "2026-01-10T00:00:00",
  Recorder: doc,
  Recorder_Type: "StandardODATA.Document_РеализацияТоваровУслуг",
  LineNumber: String(line),
  Active: true,
  AccountDr_Key: ACC,
  Сумма: 100,
});
const args = { accountKeys: [ACC], from: "2026-01-01", to: "2026-01-31" };

function replyWith(conn: ReturnType<typeof fake1C>["conn"], reply: (sorted: boolean) => object[]) {
  return vi.spyOn(conn.client, "getCollection").mockImplementation(async (url: string) => ({
    value: reply(decodeURIComponent(url).includes("$orderby")) as never[],
  }));
}

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("fetchMovements: сбой листания 1С не превращается в тихий недосчёт", () => {
  it("без повторов — один запрос, без сортировки", async () => {
    const { conn } = fake1C();
    const spy = replyWith(conn, () => [row("a"), row("b")]);
    const { rows } = await fetchMovements(conn, args);
    expect(rows).toHaveLength(2);
    expect(spy).toHaveBeenCalledTimes(1);
    expect(decodeURIComponent(String(spy.mock.calls[0]![0]))).not.toContain("$orderby");
  });

  it("повтор вместо пропущенной строки — перечитывает с устойчивой сортировкой", async () => {
    const { conn } = fake1C();
    const spy = replyWith(conn, (sorted) => (sorted ? [row("a"), row("b")] : [row("a"), row("a")]));
    const { rows, meta } = await fetchMovements(conn, args);
    expect(rows.map((r) => r["Recorder"])).toEqual(["a", "b"]);
    expect(spy).toHaveBeenCalledTimes(2);
    expect(decodeURIComponent(String(spy.mock.calls[1]![0]))).toContain(STABLE_MOVEMENT_ORDER);
    expect(meta.chunks).toBe(2);
  });

  it("повторы и с сортировкой — отказ, а не неполная сумма", async () => {
    const { conn } = fake1C();
    replyWith(conn, () => [row("a"), row("a")]);
    await expect(fetchMovements(conn, args)).rejects.toThrow("с повторами");
  });

  it("разные строки одного документа — не повтор", async () => {
    const { conn } = fake1C();
    const spy = replyWith(conn, () => [row("a", 1), row("a", 2)]);
    const { rows } = await fetchMovements(conn, args);
    expect(rows).toHaveLength(2);
    expect(spy).toHaveBeenCalledTimes(1);
  });
});
