/**
 * Обновляет снимок схемы для тестов (read-only). Запускать локально:
 *   npm run build && node --env-file=.env dist/scripts/snapshot-metadata.js [--db <имя базы>] [--out <файл>]
 *
 * Снимает $metadata базы и оставляет только объекты, которые упоминаются в src (с их табличными
 * частями, виртуальными таблицами и объявлениями функций Balance/Turnovers…), и все перечисления. Данных базы в $metadata нет — только схема;
 * перед коммитом всё равно просмотрите дифф. По умолчанию пишет test/fixtures/metadata/bp30.xml.
 */
import { readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { loadConfig } from "../config/env.js";
import { ServerContext } from "../context.js";

const arg = (name: string) => {
  const i = process.argv.indexOf(`--${name}`);
  return i >= 0 ? process.argv[i + 1] : undefined;
};

/** Имена объектов 1С, упомянутые в исходниках, — по ним фильтруется снимок. */
function referencedSets(dir: string): string[] {
  const names = new Set<string>();
  const walk = (d: string) => {
    for (const f of readdirSync(d)) {
      const p = join(d, f);
      if (statSync(p).isDirectory()) walk(p);
      else if (p.endsWith(".ts"))
        for (const m of readFileSync(p, "utf8").matchAll(
          /"((?:Catalog|Document|InformationRegister|AccumulationRegister|AccountingRegister|ChartOfAccounts|ChartOfCharacteristicTypes)_[А-Яа-яЁёA-Za-z0-9_]+)/g,
        ))
          names.add(m[1]!);
    }
  };
  walk(dir);
  return [...names];
}

export function pruneMetadata(xml: string, sets: string[]): string {
  const text = xml.replace(/\r\n/g, "\n");
  const keep = (n: string) => sets.some((s) => n === s || n.startsWith(`${s}_`));
  const blocks = (tag: string) =>
    [...text.matchAll(new RegExp(`\\t+<${tag} Name="([^"]+)"[\\s\\S]*?</${tag}>\\n`, "g"))].map(
      (m) => [m[1]!, m[0]] as const,
    );
  const entities = blocks("EntityType").filter(([n]) => keep(n));
  const complex = blocks("ComplexType").filter(([n]) => keep(n));
  const enums = blocks("EnumType");
  // Связи ссылочных полей (на какой объект ведёт Контрагент_Key и т. п.) — по имени исходного объекта.
  const associations = blocks("Association").filter(([n]) => keep(n));
  const entitySets = [...text.matchAll(/\t+<EntitySet Name="([^"]+)"[^>]*?\/>\n/g)].filter((m) =>
    keep(m[1]!),
  );
  // Функции регистров (Balance, Turnovers…) объявлены в контейнере; по типу результата видно, чьи они.
  const functions = blocks("FunctionImport").filter(([, b]) =>
    keep(/ReturnType="(?:Collection\()?StandardODATA\.([^")]+)/.exec(b)?.[1] ?? ""),
  );
  const head = text.slice(0, text.search(/\t+<EntityType /));
  const container = /\t+<EntityContainer[^>]*>\n/.exec(text)?.[0];
  if (head.length === 0 || !container)
    throw new Error("Неожиданный формат $metadata: нет EntityType/EntityContainer.");
  return (
    head +
    [...entities, ...complex, ...enums, ...associations].map(([, b]) => b).join("") +
    container +
    entitySets.map((m) => m[0]).join("") +
    functions.map(([, b]) => b).join("") +
    "\t\t\t</EntityContainer>\n\t\t</Schema>\n\t</edmx:DataServices>\n</edmx:Edmx>\n"
  );
}

async function main(): Promise<void> {
  const conn = new ServerContext(loadConfig()).db(arg("db"));
  const out = arg("out") ?? "test/fixtures/metadata/bp30.xml";
  const sets = referencedSets("src");
  const xml = pruneMetadata(await conn.client.getText("$metadata"), sets);
  writeFileSync(out, xml);
  const count = (re: RegExp) => (xml.match(re) ?? []).length;
  process.stdout.write(
    `База ${conn.cfg.name}: объектов ${count(/<EntitySet /g)}, функций ${count(/<FunctionImport /g)}, ` +
      `перечислений ${count(/<EnumType /g)}, ` +
      `${Math.round(xml.length / 1024)} КБ → ${out}\n`,
  );
}

main().catch((e: unknown) => {
  process.stderr.write(`${e instanceof Error ? e.message : String(e)}\n`);
  process.exit(1);
});
