#!/usr/bin/env node
/**
 * Совместимость с релизом конфигурации: прогоняет все инструменты против $metadata этого релиза вместо снимка из
 * репозитория и печатает, какие инструменты обращаются к объектам или полям, которых в нём нет.
 *
 *   npm run compat -- --bp30 <файл $metadata Бухгалтерии> [--ut11 <файл $metadata УТ>]
 *
 * Файл снимается с тестовой копии базы обновлённого релиза: npm run snapshot:metadata -- --db <база> --out <файл>.
 * Снимок рабочей базы содержит её доработки — в репозиторий его не кладите. Сеть и 1С не нужны: проверяется схема.
 */
import { spawnSync } from "node:child_process";
import process from "node:process";
import { existsSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

const out = (line = "") => process.stdout.write(`${line}\n`);
const err = (line) => process.stderr.write(`${line}\n`);

const arg = (name) => {
  const i = process.argv.indexOf(`--${name}`);
  return i >= 0 ? process.argv[i + 1] : undefined;
};
const env = { ...process.env };
const files = [];
for (const [flag, key] of [
  ["bp30", "ODATA_COMPAT_BP30"],
  ["ut11", "ODATA_COMPAT_UT11"],
]) {
  const file = arg(flag);
  if (!file) continue;
  if (!existsSync(file)) {
    err(`Нет файла ${file}`);
    process.exit(2);
  }
  env[key] = resolve(file);
  files.push(`${flag}: ${file}`);
}
if (!files.length) {
  err("Укажите --bp30 <файл> и/или --ut11 <файл> — $metadata проверяемого релиза.");
  process.exit(2);
}

// Только тесты, которые сверяют запросы со схемой: схемные снимки, все инструменты, запись на снимке УТ.
const tests = ["test/fake-1c.test.ts", "test/metadata-snapshot.test.ts"];
const report = join(tmpdir(), `1c-odata-compat-${process.pid}.json`);
const run = spawnSync("npx", ["vitest", "run", ...tests, "--reporter=json", `--outputFile=${report}`], {
  env,
  stdio: ["ignore", "ignore", "inherit"],
});
if (!existsSync(report)) {
  err("vitest не выдал отчёт");
  process.exit(run.status ?? 2);
}
const result = JSON.parse(readFileSync(report, "utf8"));
rmSync(report, { force: true });

out(`Проверено против: ${files.join(", ")}`);

// Сравнение схем: объекты и поля снимка из репозитория (их использует коннектор), которых в релизе нет. Часть
// инструментов при пропавшем поле не падает, а молча отвечает без него — такое видно только здесь.
const entityTypes = (xml) => {
  const types = new Map();
  for (const m of xml.matchAll(/<EntityType Name="([^"]+)">([\s\S]*?)<\/EntityType>/g))
    types.set(m[1], new Set([...m[2].matchAll(/<Property Name="([^"]+)"/g)].map((p) => p[1])));
  return types;
};
const sets = (xml) => new Set([...xml.matchAll(/<EntitySet Name="([^"]+)"/g)].map((m) => m[1]));
let lost = 0;
for (const [flag, key] of [
  ["bp30", "ODATA_COMPAT_BP30"],
  ["ut11", "ODATA_COMPAT_UT11"],
]) {
  if (!env[key]) continue;
  const base = readFileSync(
    join(import.meta.dirname, "..", "test", "fixtures", "metadata", `${flag}.xml`),
    "utf8",
  );
  const release = readFileSync(env[key], "utf8");
  const [baseTypes, releaseTypes, releaseSets] = [entityTypes(base), entityTypes(release), sets(release)];
  const lines = [];
  for (const name of sets(base)) if (!releaseSets.has(name)) lines.push(`  нет объекта ${name}`);
  for (const [name, fields] of baseTypes) {
    const now = releaseTypes.get(name);
    if (!now) continue;
    const missing = [...fields].filter((f) => !now.has(f));
    if (missing.length) lines.push(`  ${name}: нет полей ${missing.join(", ")}`);
  }
  lost += lines.length;
  out(
    `\nСхема ${flag}: ${lines.length ? "пропало по сравнению со снимком репозитория:" : "всё, что есть в снимке репозитория, на месте."}`,
  );
  if (lines.length) out(lines.join("\n"));
}
out("");
const failed = result.testResults.flatMap((f) => f.assertionResults.filter((a) => a.status === "failed"));
const passed = result.testResults.flatMap((f) =>
  f.assertionResults.filter((a) => a.status === "passed"),
).length;
for (const a of failed) {
  const why = (a.failureMessages[0] ?? "")
    .split("\n")
    .filter((l) => l.trim() && !/^\s+at /.test(l))
    .slice(0, 4)
    .map((l) => `    ${l.trim()}`)
    .join("\n");
  out(`✗ ${[...a.ancestorTitles, a.title].join(" › ")}\n${why}`);
}
out(`\nИтог: совместимы ${passed}, не совместимы ${failed.length}; пропавших объектов и полей — ${lost}.`);
process.exit(failed.length || lost ? 1 : 0);
