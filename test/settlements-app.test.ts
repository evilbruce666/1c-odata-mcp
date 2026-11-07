import { describe, expect, it } from "vitest";
import { runInNewContext } from "node:vm";
import { loadAppHtml } from "../src/mcp/apps.js";

/** Минимальный DOM: скрипт приложения выполняется по-настоящему, браузер для этого не нужен. */
function makeElement() {
  const el: Record<string, unknown> = {
    textContent: "",
    hidden: false,
    children: [] as unknown[],
    style: {},
    className: "",
    title: "",
    scrollWidth: 0,
    scrollHeight: 0,
    appendChild(c: unknown) {
      (el.children as unknown[]).push(c);
      return c;
    },
    append(...c: unknown[]) {
      (el.children as unknown[]).push(...c);
    },
    replaceChildren(...c: unknown[]) {
      el.children = c;
    },
    setAttribute() {},
    addEventListener() {},
    classList: { add() {}, remove() {}, toggle() {} },
  };
  return el;
}

function load() {
  const elements = new Map<string, ReturnType<typeof makeElement>>();
  const document = {
    documentElement: makeElement(),
    body: makeElement(),
    getElementById(id: string) {
      if (!elements.has(id)) elements.set(id, makeElement());
      return elements.get(id);
    },
    createElement: () => makeElement(),
    createTextNode: (text: string) => ({ textContent: text }),
  };
  const window = { parent: { postMessage() {} }, addEventListener() {} };
  const context: Record<string, unknown> = {
    document,
    window,
    requestAnimationFrame() {},
    console,
    Intl,
    Promise,
    setTimeout,
  };
  const script = /<script>([\s\S]*)<\/script>/.exec(loadAppHtml("settlements.html"))![1]!;
  runInNewContext(script, context);
  const render = context["render"] as (p: unknown) => void;
  const sub = () => String(elements.get("sub")!.textContent);
  return { render, sub, elements };
}

const debtors = (n: number) =>
  Array.from({ length: n }, (_, i) => ({ counterparty: `Покупатель ${i}`, ref: String(i), amount: 100 }));

describe("settlements.html на настоящих формах ответа", () => {
  it("первая страница: общее число и диапазон", () => {
    const { render, sub } = load();
    render({
      database: "ooo",
      totalReceivable: 6000,
      count: 60,
      offset: 0,
      truncated: true,
      debtors: debtors(25),
    });
    expect(sub()).toContain("60 покупателей, показаны 1–25");
  });
  it("последняя страница: диапазон, а не «10 в списке»", () => {
    const { render, sub } = load();
    render({
      database: "ooo",
      totalReceivable: 6000,
      count: 60,
      offset: 50,
      truncated: false,
      debtors: debtors(10),
    });
    expect(sub()).toContain("60 покупателей, показаны 51–60");
  });
  it("полный список без служебных полей", () => {
    const { render, sub } = load();
    render({ database: "ooo", totalReceivable: 500, count: 5, debtors: debtors(5) });
    expect(sub()).toContain("5 покупателей в списке");
    expect(sub()).not.toContain("показан");
  });
  it("кредиторы и старый ответ без count", () => {
    const { render, sub } = load();
    render({ database: "ooo", totalPayable: 200, creditors: debtors(2) });
    expect(sub()).toContain("2 поставщиков в списке");
  });
  it("срез без диапазона: неполнота названа", () => {
    const { render, sub } = load();
    render({ database: "ooo", totalReceivable: 500, count: 5, truncated: true, debtors: debtors(5) });
    expect(sub()).toContain("показана часть");
  });
  it("ответ с aging строится целиком", () => {
    const { render, elements } = load();
    render({
      database: "ooo",
      totalReceivable: 100,
      count: 1,
      aging: {
        asOf: "2026-11-01",
        overdueByDays: { "1-30": 100, "31-60": 0, "61-90": 0, "90+": 0 },
        byAge: { "0-30": 100, "31-60": 0, "61-90": 0, "90+": 0 },
        totalOverdue: 100,
      },
      debtors: [
        {
          counterparty: "Покупатель",
          ref: "1",
          amount: 100,
          debt: 100,
          advances: 0,
          overdue: 100,
          maxOverdueDays: 5,
        },
      ],
    });
    expect(elements.get("view")!.hidden).toBe(false);
  });
});
