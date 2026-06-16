/**
 * Разведчик живой базы (read-only). Запускать локально:
 *   npm run build && node --env-file=.env dist/scripts/probe-metadata.js
 *
 * Печатает в stdout: версию OData, число объектов по классам, сверку
 * ожидаемых имён БП 3.0 с фактическими и пробу виртуальных таблиц регистров.
 * Секреты не печатаются.
 */
import { loadConfig } from "../config/env.js";
import { ServerContext } from "../context.js";
import { buildQuery } from "../odata/query.js";
import { CATALOGS, DOCUMENTS, REGISTERS, COUNTERPARTY_FIELDS, resolveEntity } from "../config/mapping.js";
import type { EntityClass } from "../types/odata.js";

function line(s = ""): void {
  process.stdout.write(`${s}\n`);
}

async function main(): Promise<void> {
  const conn = new ServerContext(loadConfig()).db();
  const client = conn.client;

  line(`=== $metadata (база: ${conn.cfg.name}) ===`);
  const meta = await conn.getMetadata();
  const available = new Set(meta.entities.keys());
  line(`OData version: ${meta.odataVersion}`);
  line(`Всего EntitySet: ${meta.entities.size}`);

  const byClass = new Map<EntityClass, number>();
  for (const e of meta.entities.values()) byClass.set(e.class, (byClass.get(e.class) ?? 0) + 1);
  line("По классам:");
  for (const [cls, n] of [...byClass.entries()].sort((a, b) => b[1] - a[1])) {
    line(`  ${cls}: ${n}`);
  }

  line("\n=== Сверка ожидаемых имён БП 3.0 ===");
  const groups: Record<string, Record<string, readonly string[]>> = {
    Справочники: CATALOGS,
    Документы: DOCUMENTS,
    Регистры: REGISTERS,
  };
  for (const [groupName, group] of Object.entries(groups)) {
    line(`${groupName}:`);
    for (const [logical, candidates] of Object.entries(group)) {
      const found = resolveEntity(candidates, available);
      line(`  ${logical}: ${found ? `✓ ${found}` : `✗ не найдено (искал: ${candidates.join(", ")})`}`);
    }
  }

  line("\n=== Поля Catalog_Контрагенты ===");
  const cpSet = resolveEntity(CATALOGS.counterparties, available);
  if (cpSet) {
    const cp = meta.entities.get(cpSet);
    const names = new Set(cp?.properties.map((p) => p.name));
    for (const [k, field] of Object.entries(COUNTERPARTY_FIELDS)) {
      line(`  ${k} → ${field}: ${names.has(field) ? "✓" : "✗ нет такого поля"}`);
    }
  } else {
    line("  Catalog_Контрагенты не опубликован.");
  }

  line("\n=== Проба виртуальных таблиц регистров (/Balance) ===");
  for (const [logical, candidates] of Object.entries(REGISTERS)) {
    const reg = resolveEntity(candidates, available);
    if (!reg) {
      line(`  ${logical}: регистр не опубликован`);
      continue;
    }
    const path = `${reg}/Balance${buildQuery({ top: 1 })}`;
    try {
      const page = await client.getCollection(path);
      const sample = page.value[0];
      line(`  ${logical} (${reg}/Balance): ✓ строк в пробе: ${page.value.length}`);
      if (sample) line(`     поля: ${Object.keys(sample).join(", ")}`);
    } catch (e) {
      line(`  ${logical} (${reg}/Balance): ✗ ${e instanceof Error ? e.message : String(e)}`);
    }
  }

  line("\nГотово.");
}

main().catch((e) => {
  process.stderr.write(`Ошибка probe: ${e instanceof Error ? e.message : String(e)}\n`);
  process.exit(1);
});
