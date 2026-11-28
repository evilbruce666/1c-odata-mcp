import { readFileSync } from "node:fs";
import type { TaxRegimeKey } from "./tax-regime.js";

/**
 * Справочник сроков tax-calendar.json из корня пакета: то, чего в 1С может ещё не быть. Только сроки и статьи НК РФ —
 * суммы считает 1С.
 */
export interface TaxRule {
  id: string;
  obligation: string;
  kind: "report" | "notice" | "payment";
  title: string;
  taxpayers: Array<"ip" | "organization">;
  regimes: TaxRegimeKey[];
  period: { from: string; to: string };
  nominal: string;
  due: string;
  basis: string;
  condition?: string;
  match?: { account?: string; periodEnd?: string; reportSource?: string; kbk?: string[] };
}

export interface TaxRules {
  checkedAt: string;
  dueYears: number[];
  covers: string;
  notCovered: string[];
  /** Порог доходов для освобождения от НДС на УСН по годам дохода. */
  vatThreshold: {
    edition: string;
    inYear: { basis: string; limits: Record<string, number> };
    previousYear: { basis: string; limits: Record<string, number> };
  };
  rules: TaxRule[];
}

let cached: TaxRules | undefined;

/** Путь — от модуля, а не от текущей папки: сервер запускают из любой. */
export function taxRules(): TaxRules {
  return (cached ??= JSON.parse(
    readFileSync(new URL("../../tax-calendar.json", import.meta.url), "utf8"),
  ) as TaxRules);
}
