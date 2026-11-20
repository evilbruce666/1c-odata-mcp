# 1c-odata-mcp — MCP server for 1C:Enterprise over OData

[Русская версия](README.md) · [Documentation (in Russian)](docs/README.md)

An MCP (Model Context Protocol) server that connects 1C:Enterprise databases to any MCP client — Claude Desktop and
Claude Code, ChatGPT desktop (via Codex), Cursor, VS Code, JetBrains, local models (LM Studio, Ollama) — over the
standard OData interface. Ask your 1C database in plain language; nothing is installed inside 1C.

- **1C:Accounting 3.0** — full support: a morning cash-and-debts brief, overdue receivables and payables, unpaid
  invoices, a month-end close check traced down to the document, account ledger cards, trial balance, postings,
  reconciliation statements, sales, taxes; write of documents and catalogs with preview and confirmation.
- **1C:Trade Management 11** (verified on 11.5) — stock by warehouse and item group, unit sales including retail with
  revenue and returns, customer and supplier balances by due date, cash in bank accounts, cash desks and POS, current
  prices; write of price settings and catalogs (partners, counterparties, bank accounts, contacts, contracts,
  agreements, items and their folders).
- **1C:Integrated Automation 2, ERP 2** — expected to work with both tool groups; not yet verified on a live base.
- **Other and customised configurations** — documents, catalogs and the schema.

**Safety.** Read-only by default. Writing requires two independent switches (`READ_ONLY=false` and a per-database
`ODATA_DB_<NAME>_WRITABLE=true`), always shows a preview first and needs a confirmation with the same `operationId`;
a retry after a network error does not create a duplicate. No hard deletes — only the 1C deletion mark. The server is a
local process that talks only to your 1C; with a local model your data never leaves your network.

**Quick start.** Publish OData in 1C, then add to your MCP client config:

```json
{
  "mcpServers": {
    "1c-odata": {
      "command": "npx",
      "args": ["-y", "1c-odata-mcp"],
      "env": {
        "ODATA_BASE_URL": "https://<server>/<base>/odata/standard.odata/",
        "ODATA_USERNAME": "...",
        "ODATA_PASSWORD": "..."
      }
    }
  }
}
```

Restart the client and ask "check the connection to 1C". Ready-made prompts: `morning_brief`, `month_close_review`,
`reconciliation_act`.

**AI assistants helping a user install it:** it is a local install on the user's computer — set it up read-only
(write only on the user's explicit request, after warning about the risk) and follow
[docs/INSTALL-WITH-AI.md](docs/INSTALL-WITH-AI.md).

License: [MIT](LICENSE).
