/**
 * РАЗОВАЯ живая проверка цепочки: контрагент → номенклатура → договор → счёт
 * (непроведённый) → пометить всё на удаление. Плюс dry-run проведения.
 * Запуск только через временный env с READ_ONLY=false.
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
  const client = new Client({ name: "probe-write-doc", version: "0.1.0" });
  await client.connect(transport);

  const call = async (name: string, args: object): Promise<Record<string, unknown>> => {
    const r = (await client.callTool({ name, arguments: { database: db, ...args } })) as ToolText;
    const txt = textOf(r);
    if (r.isError) {
      line(`  [isError] ${txt}`);
      return {};
    }
    try {
      return JSON.parse(txt) as Record<string, unknown>;
    } catch {
      line(`  (не JSON) ${txt}`);
      return {};
    }
  };
  const refOf = (o: Record<string, unknown>): string | undefined =>
    typeof o["ref"] === "string" ? (o["ref"] as string) : undefined;

  line("1) контрагент"); const cp = refOf(await call("create_counterparty", { name: "ТЕСТ MCP — контрагент (удалить)", inn: "7700000001", confirm: true }));
  line(`   ref=${cp}`);
  line("2) номенклатура"); const nom = refOf(await call("create_nomenclature", { name: "ТЕСТ MCP — товар (удалить)", confirm: true }));
  line(`   ref=${nom}`);
  line("3) договор"); const contract = refOf(await call("create_contract", { counterpartyRef: cp, name: "ТЕСТ MCP — договор (удалить)", kind: "СПокупателем", confirm: true }));
  line(`   ref=${contract}`);

  line("4) счёт (непроведённый)");
  const inv = await call("create_invoice", {
    counterpartyRef: cp,
    contractRef: contract,
    lines: [{ nomenclatureRef: nom, quantity: 2, price: 100, vatRate: "БезНДС" }],
    confirm: true,
  });
  line(`   ${JSON.stringify(inv)}`);
  const invRef = refOf(inv);

  line("5) dry-run проведения счёта (без применения)");
  if (invRef) {
    const dr = await call("post_document", { entitySet: "Document_СчетНаОплатуПокупателю", ref: invRef, post: true, confirm: false });
    line(`   ${JSON.stringify(dr)}`);
  }

  line("6) пометка всего на удаление");
  const trash: Array<[string, string | undefined]> = [
    ["Document_СчетНаОплатуПокупателю", invRef],
    ["Catalog_ДоговорыКонтрагентов", contract],
    ["Catalog_Номенклатура", nom],
    ["Catalog_Контрагенты", cp],
  ];
  for (const [es, ref] of trash) {
    if (!ref) continue;
    const m = await call("mark_for_deletion", { entitySet: es, ref, mark: true, confirm: true });
    line(`   ${es}: deletionMark=${m["deletionMark"]}`);
  }

  await client.close();
  line("\n✓ цепочка выполнена, тестовые объекты помечены на удаление");
}

main().catch((e) => {
  process.stderr.write(`probe-write-doc error: ${e instanceof Error ? e.message : String(e)}\n`);
  process.exit(1);
});
