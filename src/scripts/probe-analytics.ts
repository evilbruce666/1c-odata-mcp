/**
 * Проверка аналитических путей (дебиторка / остатки) на живых данных.
 * Запуск: node --env-file=.env dist/scripts/probe-analytics.js
 * Использует те же хелперы, что и инструменты read.analytics.get_debtors / read.analytics.get_inventory.
 */
import { loadConfig } from "../config/env.js";
import { ServerContext } from "../context.js";
import { ACCOUNT_PREFIX, CATALOGS, resolveEntity } from "../config/mapping.js";
import { balanceByAccounts, resolveAccounts, resolveNames, num } from "../odata/accounting.js";

function line(s = ""): void {
  process.stdout.write(`${s}\n`);
}

async function main(): Promise<void> {
  const conn = new ServerContext(loadConfig()).db();
  line(`база: ${conn.cfg.name}`);

  line("=== ДЕБИТОРКА (сч. 62) ===");
  const recAccts = await resolveAccounts(conn, ACCOUNT_PREFIX.receivables);
  line(
    `счетов: ${recAccts.length} (${recAccts
      .slice(0, 4)
      .map((a) => a.code)
      .join(", ")}…)`,
  );
  const recRows = await balanceByAccounts(
    conn,
    recAccts.map((a) => a.key),
  );
  line(`строк сальдо: ${recRows.length}`);
  const byCp = new Map<string, number>();
  for (const r of recRows) {
    const cp = String(r["ExtDimension1"] ?? "");
    if (!cp) continue;
    byCp.set(cp, (byCp.get(cp) ?? 0) + num(r["СуммаBalanceDr"]) - num(r["СуммаBalanceCr"]));
  }
  const cpSet = resolveEntity(CATALOGS.counterparties, await conn.available());
  const cpNames = cpSet ? await resolveNames(conn, cpSet, byCp.keys()) : new Map();
  const debtors = [...byCp.entries()]
    .filter(([, v]) => v > 0.005)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  line(`должников: ${[...byCp.values()].filter((v) => v > 0.005).length}, топ-5:`);
  for (const [ref, amt] of debtors) line(`  ${cpNames.get(ref) ?? ref}: ${amt.toFixed(2)}`);

  line("\n=== ОСТАТКИ ТОВАРОВ (сч. 41/10/43) ===");
  const invAccts = await resolveAccounts(conn, ACCOUNT_PREFIX.inventory);
  line(`счетов: ${invAccts.length}`);
  const invRows = await balanceByAccounts(
    conn,
    invAccts.map((a) => a.key),
  );
  line(`строк сальдо: ${invRows.length}`);
  const byItem = new Map<string, { qty: number; amount: number }>();
  for (const r of invRows) {
    const it = String(r["ExtDimension1"] ?? "");
    if (!it) continue;
    const c = byItem.get(it) ?? { qty: 0, amount: 0 };
    c.qty += num(r["КоличествоBalanceDr"]) - num(r["КоличествоBalanceCr"]);
    c.amount += num(r["СуммаBalanceDr"]) - num(r["СуммаBalanceCr"]);
    byItem.set(it, c);
  }
  const nomSet = resolveEntity(CATALOGS.nomenclature, await conn.available());
  const nomNames = nomSet ? await resolveNames(conn, nomSet, byItem.keys()) : new Map();
  const items = [...byItem.entries()].sort((a, b) => b[1].amount - a[1].amount).slice(0, 5);
  line(`позиций: ${byItem.size}, топ-5 по сумме:`);
  for (const [ref, v] of items)
    line(`  ${nomNames.get(ref) ?? ref}: кол-во ${v.qty}, сумма ${v.amount.toFixed(2)}`);

  line("\nГотово.");
}

main().catch((e) => {
  process.stderr.write(`Ошибка: ${e instanceof Error ? e.message : String(e)}\n`);
  process.exit(1);
});
