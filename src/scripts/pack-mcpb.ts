/**
 * Собирает .mcpb для Claude Desktop: чистая установка рабочих зависимостей, манифест в формате MCPB, zip.
 * Запуск после сборки: npm run bundle -- [путь к файлу]. Подробности — docs/BUILDING-MCPB.md.
 */
import { execFileSync } from "node:child_process";
import {
  copyFileSync,
  cpSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { bundleManifest } from "../mcp/bundle-manifest.js";

/** В рантайме не нужны: типы, sourcemaps, исходники .ts и markdown пакетов. */
const UNNEEDED = /(\.map|\.d\.ts|\.d\.mts|\.d\.cts)$/;
const UNNEEDED_IN_MODULES = /(\.ts|\.md|^changelog.*)$/i;

function prune(dir: string, inModules = false): void {
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) prune(path, inModules || name === "node_modules");
    else if (UNNEEDED.test(name) || (inModules && UNNEEDED_IN_MODULES.test(name))) rmSync(path);
  }
}

const pkg = JSON.parse(readFileSync("package.json", "utf8")) as { version: string };
const out = resolve(process.argv[2] ?? join(tmpdir(), `1c-odata-mcp-${pkg.version}.mcpb`));
const stage = mkdtempSync(join(tmpdir(), "1c-odata-mcpb-"));
try {
  cpSync("dist", join(stage, "dist"), { recursive: true });
  // Приложения MCP (долги по срокам) сервер читает из apps/ рядом с dist.
  cpSync("apps", join(stage, "apps"), { recursive: true });
  rmSync(join(stage, "dist", "scripts"), { recursive: true, force: true });
  for (const file of ["package.json", "package-lock.json", "README.md", "LICENSE"]) {
    copyFileSync(file, join(stage, file));
  }
  const manifest = JSON.parse(readFileSync("manifest.json", "utf8")) as Parameters<typeof bundleManifest>[0];
  writeFileSync(join(stage, "manifest.json"), JSON.stringify(bundleManifest(manifest), null, 2));
  execFileSync("npm", ["ci", "--omit=dev", "--ignore-scripts"], { cwd: stage, stdio: "ignore" });
  prune(stage);
  rmSync(out, { force: true });
  execFileSync("zip", ["-r", "-X", "-q", "-9", out, ".", "-x", ".*"], { cwd: stage });
  process.stdout.write(`${out}: ${statSync(out).size} байт\n`);
} finally {
  rmSync(stage, { recursive: true, force: true });
}
