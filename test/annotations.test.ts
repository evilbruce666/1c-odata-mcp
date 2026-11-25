import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

/**
 * Подсказки клиенту (MCP): destructiveHint — инструмент перезаписывает или удаляет имеющееся (клиент спрашивает
 * подтверждение), idempotentHint — повтор тех же аргументов ничего не добавляет. Раньше destructive были только
 * проведение и пометка, и клиент считал замену всех строк или update_entity добавляющими (аудит 25.11).
 * Списки — по коду инструментов; новый пишущий инструмент надо отнести сюда осознанно.
 */
const manifest = JSON.parse(readFileSync("manifest.json", "utf8")) as {
  tools: Array<{ name: string; annotations: Record<string, boolean>; inputSchema: { properties?: object } }>;
};
/** makeMain заменяет у владельца прежний основной счёт или контакт — это перезапись (ревью 25.11). */
const REPLACES_MAIN = manifest.tools
  .filter((t) => "makeMain" in (t.inputSchema.properties ?? {}))
  .map((t) => t.name);
const hinted = (key: string) =>
  manifest.tools
    .filter((t) => t.annotations?.[key] === true)
    .map((t) => t.name)
    .sort();

const DESTRUCTIVE = [
  "write.accounting.set_contract_settlement_settings",
  "write.accounting.set_settlement_settings",
  "write.catalog.update_nomenclature",
  "write.counterparty.create_bank_account",
  "write.counterparty.create_contact_person",
  "write.counterparty.set_group",
  "write.counterparty.update_counterparty",
  "write.document.align_bank_payment_amounts",
  "write.document.apply_bank_settlement_settings",
  "write.document.post_document",
  "write.document.remove_document_line",
  "write.document.update_document_lines",
  "write.entity.mark_for_deletion",
  "write.entity.move_to_folder",
  "write.entity.update_entity",
];
const IDEMPOTENT = [
  "write.catalog.update_nomenclature",
  "write.counterparty.update_counterparty",
  "write.document.update_document_lines",
  "write.entity.mark_for_deletion",
  "write.entity.move_to_folder",
  "write.entity.update_entity",
];

describe("аннотации инструментов", () => {
  it("destructiveHint — ровно у перезаписывающих", () => {
    expect(hinted("destructiveHint")).toEqual(DESTRUCTIVE);
  });
  it("инструмент с makeMain — destructive: он заменяет основной у владельца", () => {
    expect(REPLACES_MAIN.length).toBeGreaterThanOrEqual(2);
    expect(REPLACES_MAIN.filter((n) => !hinted("destructiveHint").includes(n))).toEqual([]);
  });
  it("idempotentHint — ровно у правок, где повтор ничего не добавляет", () => {
    expect(hinted("idempotentHint")).toEqual(IDEMPOTENT);
  });
  it("чтения — только readOnlyHint; создания и добавление строки не destructive и не idempotent", () => {
    for (const t of manifest.tools) {
      if (t.name.startsWith("read."))
        expect(t.annotations, t.name).toEqual({ readOnlyHint: true, openWorldHint: true });
      if (/\.create_|add_document_line|copy_document/.test(t.name) && !REPLACES_MAIN.includes(t.name)) {
        expect(t.annotations["destructiveHint"], t.name).toBe(false);
        expect(t.annotations["idempotentHint"], t.name).toBeUndefined();
      }
    }
  });
});
