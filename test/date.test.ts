import { describe, it, expect } from "vitest";
import { odataDate } from "../src/tools/write.js";

describe("odataDate", () => {
  // Регресс: дата НЕ должна уезжать на день назад в плюсовых поясах.
  // Раньше форматирование шло через toISOString() (UTC), и полночь 15-го в MSK (+3)
  // превращалась в 21:00 14-го UTC → 1С записывала 14-е вместо 15-го.
  it("сохраняет календарный день (локальная полночь, без UTC-сдвига)", () => {
    const d = new Date(2026, 5, 15, 0, 0, 0); // локально 15 июня 00:00
    expect(odataDate(d)).toBe("2026-06-15T00:00:00");
  });

  it("формат Edm.DateTime без зоны", () => {
    expect(odataDate(new Date(2026, 0, 2, 9, 8, 7))).toBe("2026-01-02T09:08:07");
  });
});
