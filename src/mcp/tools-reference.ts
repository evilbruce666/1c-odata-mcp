/**
 * Справочник параметров инструментов (docs/TOOLS-REFERENCE.md) — из manifest.json, то есть из того же tools/list,
 * что видит MCP-клиент. Файл собирается `npm run manifest`, тест сверяет его с кодом: справочник не устаревает.
 */

interface JsonSchema {
  type?: string | string[];
  enum?: unknown[];
  items?: JsonSchema;
  properties?: Record<string, JsonSchema>;
  required?: string[];
  description?: string;
  default?: unknown;
  anyOf?: JsonSchema[];
  minimum?: number;
  maximum?: number;
}

export interface ManifestTool {
  name: string;
  title?: string;
  description?: string;
  inputSchema?: JsonSchema;
}

/** Ячейка таблицы: без переводов строк и с экранированной чертой. */
const cell = (s: string) =>
  s
    .replace(/\s*\n\s*/g, " ")
    .replace(/\\/g, "\\\\")
    .replace(/\|/g, "\\|");

function typeOf(s: JsonSchema): string {
  if (s.enum) return s.enum.map((v) => `\`${String(v)}\``).join(", ");
  if (s.anyOf) return s.anyOf.map(typeOf).join(" или ");
  const t = Array.isArray(s.type) ? s.type.join(" или ") : (s.type ?? "любое");
  if (t === "array" && s.items) {
    const inner = s.items.properties
      ? `объект: ${Object.keys(s.items.properties).join(", ")}`
      : typeOf(s.items);
    return `список (${inner})`;
  }
  if (t === "object" && s.properties) return `объект: ${Object.keys(s.properties).join(", ")}`;
  return t;
}

function paramRows(schema: JsonSchema | undefined): string[] {
  const props = schema?.properties ?? {};
  const required = new Set(schema?.required ?? []);
  return Object.entries(props).map(([name, p]) => {
    const extra = [
      p.default !== undefined ? `по умолчанию \`${JSON.stringify(p.default)}\`` : "",
      // Верхняя граница вида Number.MAX_SAFE_INTEGER — техническая, не для читателя.
      ...(() => {
        const max = p.maximum !== undefined && p.maximum < 1e12 ? p.maximum : undefined;
        const min = p.minimum !== undefined && p.minimum > -1e12 ? p.minimum : undefined;
        if (min === undefined && max === undefined) return [""];
        if (max === undefined) return [`не меньше ${min}`];
        return [min === undefined ? `не больше ${max}` : `от ${min} до ${max}`];
      })(),
    ]
      .filter(Boolean)
      .join("; ");
    const description = [p.description ?? "", extra].filter(Boolean).join(". ");
    return `| \`${name}\` | ${cell(typeOf(p))} | ${required.has(name) ? "да" : ""} | ${cell(description)} |`;
  });
}

export function renderToolsReference(tools: readonly ManifestTool[]): string {
  const out: string[] = [
    "[← Документация](README.md)",
    "",
    "# Справочник параметров инструментов",
    "",
    "Собирается из кода (`npm run manifest`) — тот же список, что видит MCP-клиент. Не правьте вручную: тест сверяет " +
      "файл с кодом. Обзор инструментов по конфигурациям — [TOOLS.md](TOOLS.md).",
    "",
    `Всего ${tools.length}. У инструментов записи \`confirm\` и \`operationId\` — см. [WRITE.md](WRITE.md).`,
  ];
  let group = "";
  for (const t of [...tools].sort((a, b) => a.name.localeCompare(b.name))) {
    const g = t.name.split(".").slice(0, 2).join(".");
    if (g !== group) {
      group = g;
      out.push("", `## ${g}`);
    }
    out.push("", `### \`${t.name}\`${t.title ? ` — ${t.title}` : ""}`, "", (t.description ?? "").trim());
    const rows = paramRows(t.inputSchema);
    out.push(
      "",
      ...(rows.length
        ? ["| Параметр | Тип | Обязательный | Описание |", "|---|---|---|---|", ...rows]
        : ["Без параметров."]),
    );
  }
  return `${out.join("\n")}\n`;
}
