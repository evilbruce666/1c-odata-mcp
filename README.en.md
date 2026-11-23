# 1c-odata-mcp — MCP server for 1C:Enterprise over OData

[Русская версия](README.md) · [Documentation (in Russian)](docs/README.md)

An MCP (Model Context Protocol) server that connects 1C:Enterprise databases to any MCP client — Claude Desktop and
Claude Code, ChatGPT desktop (via Codex), Cursor, VS Code, JetBrains, local models (LM Studio, Ollama) — over the
standard OData interface. Ask your 1C database in plain language; nothing is installed inside 1C.

- **1C:Accounting 3.0** — full support: a morning cash-and-debts brief, overdue receivables and payables, unpaid
  invoices, a month-end close check traced down to the document, account ledger cards, trial balance, postings,
  reconciliation statements, revenue and gross margin by item group and period with year-over-year comparison, sales, taxes; write of documents and catalogs with preview and confirmation, including payments split across contracts, transfers between own accounts and card acquiring receipts.
- **1C:Trade Management 11** (verified on 11.5) — stock by warehouse and item group (with cost from the «Себестоимость товаров» register), unit sales and gross margin in any breakdown (store, department, manager, customer, item group, period) with year-over-year comparison, including retail with
  revenue and returns, customer and supplier balances by due date, cash in bank accounts, cash desks and POS, current
  prices; write of price settings, catalogs (partners, counterparties, bank accounts, contacts, contracts,
  agreements, items and their folders), warehouse documents, cash and bank documents, payment requests, sales and
  purchases (orders, invoices, shipments, purchases, returns, VAT invoices), with posting checked against the stock,
  cash and settlement registers. Per-tool evidence — [docs/SUPPORT.md](docs/SUPPORT.md) (in Russian).
- **1C:Integrated Automation 2, ERP 2** — expected to work with both tool groups; not yet verified on a live base and
  not part of the full-support claim.
- **Other and customised configurations** — documents, catalogs and the schema.

**Safety.** Read-only by default. Writing requires two independent switches (`READ_ONLY=false` and a per-database
`ODATA_WRITABLE=true` for a single database or `ODATA_DB_<NAME>_WRITABLE=true` for each of several), always shows a preview first and needs a confirmation with the same `operationId`;
a retry after a network error does not create a duplicate. No hard deletes — only the 1C deletion mark. The server is a
local process that talks only to your 1C; with a local model your data never leaves your network.

**Server install (bot, own app).** `MCP_TRANSPORT=http` with a mandatory token: read-only by default. Writing over
HTTP runs as a separate instance and only when all of these hold: `READ_ONLY=false`, a writable database, an explicit
write journal on a volume (`ODATA_WRITE_JOURNAL_DIR`, checked at start), and the write tools plus
`write.operation.status` listed by name in `TOOLS_ALLOW`. Who confirms a write is up to the integration — the server
cannot tell a person from a program; it guarantees preview, no duplicates and a readback from 1C. Details and a compose
example — [docs/CONNECTING.md](docs/CONNECTING.md#запись-по-http) (in Russian).

**Versioning.** Semantic versioning from 1.0.0: tool names, parameters, response fields, error kinds and environment
variables change incompatibly only in a major version — [docs/VERSIONING.md](docs/VERSIONING.md) (in Russian).

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
        "ODATA_PASSWORD": "...",
        "TOOLS_PROFILE": "read"
      }
    }
  }
}
```

`TOOLS_PROFILE=read` exposes only the read tools: less for the model to read, and it cannot write to 1C. Remove
this line to enable writing.

Restart the client and ask "check the connection to 1C". Ready-made prompts: `morning_brief`, `sales_review`
(revenue and margin for a month against the previous month and year — in both configurations), `month_close_review`,
`reconciliation_act`.

**Updates.** With `npx` the connector updates itself when the client starts — just restart it. A global install
(`npm install -g`) does not update and takes over `npx`; with writing enabled, pin the version
(`1c-odata-mcp@1.1.0`). Details — [docs/CONNECTING.md](docs/CONNECTING.md#обновление) (in Russian).

**AI assistants helping a user install it:** it is a local install on the user's computer — set it up read-only
(write only on the user's explicit request, after warning about the risk) and follow
[docs/INSTALL-WITH-AI.md](docs/INSTALL-WITH-AI.md).

License: [MIT](LICENSE).
