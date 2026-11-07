import { describe, expect, it } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { parseConfig } from "../src/config/env.js";
import { createServer } from "../src/mcp/server.js";
import { PROMPTS, promptTools } from "../src/mcp/prompts.js";
import { fake1C } from "./helpers/fake-1c.js";

async function connect(env: Record<string, string>) {
  const fake = fake1C({ READ_ONLY: "true", ...env });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(fake.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  await client.connect(b);
  const names = (await client.listTools()).tools.map((t) => t.name);
  return { client, names };
}

describe("фильтр инструментов", () => {
  it("по умолчанию видны все, и при READ_ONLY=true тоже: предпросмотр записи остаётся", async () => {
    const { names, client } = await connect({});
    expect(names).toHaveLength(76);
    expect(names.filter((n) => n.startsWith("write."))).toHaveLength(42);
    expect((await client.listPrompts()).prompts).toHaveLength(PROMPTS.length);
    expect(client.getInstructions()).not.toContain("TOOLS_");
  });

  it("TOOLS_PROFILE=read: только чтение, запись нельзя вызвать", async () => {
    const { names, client } = await connect({ TOOLS_PROFILE: "read" });
    expect(names).toHaveLength(34);
    expect(names.every((n) => n.startsWith("read."))).toBe(true);
    expect(client.getInstructions()).toContain("TOOLS_PROFILE=read");
    await expect(client.callTool({ name: "write.sales.create_invoice", arguments: {} })).rejects.toThrow(
      /not found/,
    );
    expect((await client.listPrompts()).prompts).toHaveLength(PROMPTS.length);
  });

  it("TOOLS_ALLOW: точный список, другой список даёт другой набор", async () => {
    const one = await connect({ TOOLS_ALLOW: "read.system.list_databases, read.analytics.get_debtors" });
    expect(one.names.sort()).toEqual(["read.analytics.get_debtors", "read.system.list_databases"]);
    expect(one.client.getInstructions()).toContain("TOOLS_ALLOW");
    const two = await connect({ TOOLS_ALLOW: "read.system.list_databases" });
    expect(two.names).toEqual(["read.system.list_databases"]);
  });

  it("сценарий показывается, только если доступны все его инструменты", async () => {
    const brief = PROMPTS.find((p) => p.name === "morning_brief")!;
    expect(promptTools(brief)).toContain("read.analytics.get_debtors");
    const partial = await connect({ TOOLS_ALLOW: promptTools(brief).slice(1).join(",") });
    expect((await partial.client.listPrompts().catch(() => ({ prompts: [] }))).prompts).toEqual([]);
    const full = await connect({ TOOLS_ALLOW: promptTools(brief).join(",") });
    expect((await full.client.listPrompts()).prompts.map((p) => p.name)).toEqual(["morning_brief"]);
  });

  it("опечатка в TOOLS_ALLOW — отказ при создании сервера с именем", () => {
    const fake = fake1C({ TOOLS_ALLOW: "read.system.list_databases,read.nope.none" });
    expect(() => createServer(fake.ctx)).toThrow(/read\.nope\.none/);
  });

  describe("разбор настройки", () => {
    const base = { ODATA_BASE_URL: "http://1c.test/", ODATA_USERNAME: "u", ODATA_PASSWORD: "p" };
    it("вместе не задаются", () => {
      expect(() => parseConfig({ ...base, TOOLS_PROFILE: "read", TOOLS_ALLOW: "read.a.b" })).toThrow(
        /TOOLS_PROFILE и TOOLS_ALLOW/,
      );
    });
    it("неизвестный профиль — отказ", () => {
      expect(() => parseConfig({ ...base, TOOLS_PROFILE: "minimal" })).toThrow(/TOOLS_PROFILE/);
    });
    it("пустые значения из шаблона .env — не заданы", () => {
      const { behavior } = parseConfig({ ...base, TOOLS_PROFILE: "", TOOLS_ALLOW: " , " });
      expect(behavior.toolsProfile).toBe("all");
      expect(behavior.toolsAllow).toBeUndefined();
    });
  });
});
