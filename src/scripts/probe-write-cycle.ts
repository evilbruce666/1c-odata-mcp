/**
 * РАЗОВАЯ живая проверка цикла: поставщик+договор+поступление, покупатель+договор+реализация
 * (оба непроведённые) → чтение назад → пометка на удаление. Без проведения.
 * Запуск через временный env с READ_ONLY=false.
 */
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

type ToolText = { content: Array<{ type: string; text?: string }>; isError?: boolean };
const line = (s = ""): void => void process.stdout.write(`${s}\n`);
const textOf = (r: ToolText): string => r.content.map((c) => c.text ?? "").join("\n");

async function main(): Promise<void> {
  const db = process.env.PROBE_DB ?? "ip";
  const env: Record<string, string> = {};
  for (const [k, v] of Object.entries(process.env)) if (v !== undefined) env[k] = v;
  const transport = new StdioClientTransport({ command: "node", args: ["dist/index.js"], env });
  const client = new Client({ name: "probe-write-cycle", version: "0.1.0" });
  await client.connect(transport);

  const call = async (name: string, args: object): Promise<Record<string, unknown>> => {
    const r = (await client.callTool({ name, arguments: { database: db, ...args } })) as ToolText;
    const txt = textOf(r);
    if (r.isError) { line(`  [isError] ${txt}`); return {}; }
    try { return JSON.parse(txt) as Record<string, unknown>; } catch { return { raw: txt }; }
  };
  const ref = (o: Record<string, unknown>): string | undefined => (typeof o["ref"] === "string" ? (o["ref"] as string) : undefined);

  const nom = ref(await call("create_nomenclature", { name: "ТЕСТ MCP — товар цикла (удалить)", confirm: true }));
  line(`номенклатура: ${nom}`);

  const sup = ref(await call("create_counterparty", { name: "ТЕСТ MCP — поставщик (удалить)", inn: "7700000002", confirm: true }));
  const supC = ref(await call("create_contract", { counterpartyRef: sup, name: "ТЕСТ MCP — договор поставки", kind: "СПоставщиком", confirm: true }));
  const purchase = ref(await call("create_purchase", { counterpartyRef: sup, contractRef: supC, lines: [{ nomenclatureRef: nom, quantity: 10, price: 50 }], confirm: true }));
  line(`поступление: ${purchase}`);

  const cust = ref(await call("create_counterparty", { name: "ТЕСТ MCP — покупатель (удалить)", inn: "7700000003", confirm: true }));
  const custC = ref(await call("create_contract", { counterpartyRef: cust, name: "ТЕСТ MCP — договор продажи", kind: "СПокупателем", confirm: true }));
  const shipment = ref(await call("create_shipment", { counterpartyRef: cust, contractRef: custC, lines: [{ nomenclatureRef: nom, quantity: 3, price: 100 }], confirm: true }));
  line(`реализация: ${shipment}`);

  // Читаем документы назад через get_document
  for (const [es, r, label] of [
    ["Document_ПоступлениеТоваровУслуг", purchase, "Поступление"],
    ["Document_РеализацияТоваровУслуг", shipment, "Реализация"],
  ] as Array<[string, string | undefined, string]>) {
    if (!r) continue;
    const d = await call("get_document", { entitySet: es, ref: r });
    const t = (d["Товары"] as Array<Record<string, unknown>>) ?? [];
    line(`${label}: Posted=${d["Posted"]} СуммаДокумента=${d["СуммаДокумента"]} строк=${t.length}` +
      (t[0] ? ` (Кол-во ${t[0]["Количество"]} × Цена ${t[0]["Цена"]} = ${t[0]["Сумма"]})` : ""));
  }

  line("пометка на удаление…");
  const trash: Array<[string, string | undefined]> = [
    ["Document_РеализацияТоваровУслуг", shipment],
    ["Document_ПоступлениеТоваровУслуг", purchase],
    ["Catalog_ДоговорыКонтрагентов", custC],
    ["Catalog_ДоговорыКонтрагентов", supC],
    ["Catalog_Контрагенты", cust],
    ["Catalog_Контрагенты", sup],
    ["Catalog_Номенклатура", nom],
  ];
  for (const [es, r] of trash) if (r) await call("mark_for_deletion", { entitySet: es, ref: r, mark: true, confirm: true });
  line("готово (всё помечено на удаление)");

  await client.close();
}

main().catch((e) => {
  process.stderr.write(`probe-write-cycle error: ${e instanceof Error ? e.message : String(e)}\n`);
  process.exit(1);
});
