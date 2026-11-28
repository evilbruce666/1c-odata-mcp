import { createServer } from "../../src/mcp/server.js";
import { fake1C, type FakeOptions, type FakeQuery } from "./fake-1c.js";

type Result = { isError?: boolean; content?: unknown[]; structuredContent?: Record<string, unknown> };
type Registered = {
  inputSchema?: { parse: (a: unknown) => Record<string, unknown> };
  handler: (a: Record<string, unknown>, e: unknown) => Promise<Result>;
};

/**
 * Один вызов инструмента против поддельной 1С так же, как делает SDK (аргументы проходят схему). data — свои строки
 * объекта; queries — GET-запросы коллекций, дошедшие до данных.
 */
export async function callTool(
  name: string,
  raw: Record<string, unknown>,
  data?: FakeOptions["data"],
  schema?: FakeOptions["schema"],
) {
  const queries: FakeQuery[] = [];
  const fake = fake1C(
    { READ_ONLY: "true" },
    {
      ...(schema ? { schema } : {}),
      data: (q) => {
        queries.push(q);
        return data?.(q);
      },
    },
  );
  const tool = (createServer(fake.ctx) as unknown as { _registeredTools: Record<string, Registered> })
    ._registeredTools[name]!;
  const result = await tool.handler(tool.inputSchema ? tool.inputSchema.parse(raw) : raw, {});
  return { result, queries, fake, violations: fake.violations };
}
