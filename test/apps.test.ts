import { describe, expect, it } from "vitest";
import { APPS, APP_MIME_TYPE, appMetaFor, loadAppHtml } from "../src/mcp/apps.js";

describe("MCP Apps", () => {
  it("инструмент с приложением получает ссылку на ui://-ресурс, остальные — нет", () => {
    expect(appMetaFor("read.analytics.get_debtors")).toEqual({
      ui: { resourceUri: "ui://1c-odata/settlements.html" },
      "ui/resourceUri": "ui://1c-odata/settlements.html",
    });
    expect(appMetaFor("read.analytics.get_creditors")).toBeDefined();
    expect(appMetaFor("read.document.search_documents")).toBeUndefined();
    expect(APP_MIME_TYPE).toBe("text/html;profile=mcp-app");
  });

  for (const app of APPS) {
    describe(app.file, () => {
      const html = loadAppHtml(app.file);

      it("самодостаточный: ничего не загружает извне", () => {
        expect(html).not.toMatch(/<(script|link|img|iframe)[^>]+(src|href)\s*=/i);
        expect(html).not.toMatch(/@import|url\(\s*["']?https?:/i);
        expect(html).not.toMatch(/\bfetch\(|XMLHttpRequest|WebSocket/);
      });

      it("строки из 1С не попадают в разметку: нет innerHTML и подобного", () => {
        expect(html).not.toMatch(
          /innerHTML|outerHTML|insertAdjacentHTML|document\.write|\beval\(|new Function/,
        );
      });

      it("отвечает на протокол MCP Apps", () => {
        expect(html).toContain("ui/initialize");
        expect(html).toContain("ui/notifications/tool-result");
        expect(html).toContain("ui/notifications/size-changed");
      });
    });
  }
});

describe("settlements.html: неполный список", () => {
  it("признак берётся из аргумента render, а не из несуществующей переменной", () => {
    const html = loadAppHtml("settlements.html");
    const body = /function render\(p\) \{[\s\S]*?\n\}/.exec(html)![0];
    expect(body).toContain("p.truncated");
    expect(body).toContain("p.count");
    expect(body).not.toMatch(/\bdata\.truncated\b/);
  });
});
