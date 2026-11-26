#!/usr/bin/env node
/**
 * Эталон публичного API для test/api-contract.test.ts — инструменты и сценарии выпускаемой версии.
 * Запускается при выпуске, после подъёма версии и `npm run manifest`. Сначала сверяет текущий API с прежним
 * эталоном: если совместимость нарушена, эталон не обновляется — иначе поломку спрятало бы само обновление.
 */
import { spawnSync } from "node:child_process";
import { readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import process from "node:process";

const root = join(import.meta.dirname, "..");
const baselineFile = join(root, "test/fixtures/api-contract/release.json");
const previous = JSON.parse(readFileSync(baselineFile, "utf8")).version;
const next = JSON.parse(readFileSync(join(root, "manifest.json"), "utf8")).version;
if (previous === next) {
  process.stderr.write(
    `Версия в manifest.json (${next}) та же, что у эталона: эталон — это выпущенный API. Поднимите версию и запустите npm run manifest.\n`,
  );
  process.exit(1);
}
const check = spawnSync("npx", ["vitest", "run", "test/api-contract.test.ts", "-t", "ничего не удалено"], {
  cwd: root,
  stdio: "inherit",
});
if (check.status !== 0) {
  process.stderr.write("\nAPI несовместим с прежним выпуском — эталон не обновлён (docs/VERSIONING.md).\n");
  process.exit(1);
}
const manifest = JSON.parse(readFileSync(join(root, "manifest.json"), "utf8"));
const { version, tools, prompts } = manifest;
writeFileSync(
  join(root, "test/fixtures/api-contract/release.json"),
  JSON.stringify({ version, tools, prompts }),
);
process.stdout.write(
  `Эталон API — версия ${version}: инструментов ${tools.length}, сценариев ${prompts.length}.\n`,
);
