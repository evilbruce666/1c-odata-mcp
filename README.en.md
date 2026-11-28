# 1c-odata-mcp — MCP server for 1C:Enterprise over OData

[Русская версия](README.md) · [Documentation (in Russian)](docs/README.md)

[![Glama](https://glama.ai/mcp/servers/evilbruce666/1c-odata-mcp/badges/score.svg)](https://glama.ai/mcp/servers/evilbruce666/1c-odata-mcp)

**Connect AI to your 1C database without changing its configuration.** This is an MCP server that runs on your own
computer and lets Claude, ChatGPT desktop (via Codex), Cursor or a local model work with a 1C:Enterprise database
through the standard OData interface. You ask in plain language — who owes us, stock by store, is August ready to
close — and get the numbers from your own base.

Nothing goes into 1C itself: no extensions, no COM connection, no SQL access, just a published OData endpoint. Most of
the documentation is in Russian, because most users are 1C people in Russia, Belarus, Kazakhstan and Kyrgyzstan; if
you don't read Russian, your assistant can read it for you.

## What it covers

- **1C:Accounting 3.0** — full support: a morning cash-and-debts brief, overdue receivables and payables, unpaid
  invoices, a month-end close check traced down to the document, account ledger cards, trial balance, postings,
  reconciliation statements, revenue and gross margin by item group and period with year-over-year comparison, sales,
  taxes; write of documents and catalogs with preview and confirmation, including payments split across contracts,
  transfers between own accounts and card acquiring receipts. Also 1C:Accounting for Belarus through the
  chart-of-accounts settings.
- **1C:Trade Management 11** (verified on 11.5) — stock by warehouse and item group (with cost from the «Себестоимость
  товаров» register), unit sales and gross margin in any breakdown (store, department, manager, customer, item group,
  period) with year-over-year comparison, including retail with revenue and returns, customer and supplier balances by
  due date, cash in bank accounts, cash desks and POS, current prices; write of price settings, catalogs, warehouse,
  cash and bank documents, payment requests, sales and purchases, with posting checked against the stock, cash and
  settlement registers. Per-tool evidence — [docs/SUPPORT.md](docs/SUPPORT.md).
- **1C:Integrated Automation 2, ERP 2** — expected to work with both tool groups; not yet verified on a live base and
  not part of the full-support claim.
- **Other and customised configurations** (UNF, ZUP, Retail and others) — documents, catalogs and the schema; no
  dedicated analytics for them.

93 tools (45 read/analytics + 48 write), each with an `outputSchema`; names are `<read|write>.<category>.<name>`.
Reference — [docs/TOOLS.md](docs/TOOLS.md), all parameters — [docs/TOOLS-REFERENCE.md](docs/TOOLS-REFERENCE.md).

## Quick start

1. **Node.js 22 or newer** (`node -v`).
2. **Publish OData** in 1C and add the objects you need to its «Состав» (content) — your 1C administrator or cloud
   provider does this; step by step — [docs/ODATA-SETUP.md](docs/ODATA-SETUP.md).
3. **Add the server** to your MCP client config:

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
this line to enable writing. On Windows, if the client cannot start `npx`, use
`"command": "cmd", "args": ["/c", "npx", "-y", "1c-odata-mcp"]`.

4. **Restart the client** and ask "check the connection to 1C" — a healthy answer names the OData version and the
   number of published objects.

If it does not work, check from a terminal: `npx -y 1c-odata-mcp --check --config "<path to the client config>"` reads
the `env` block of that file (no need to retype the password), makes one `$metadata` request per database and tells you
whether 1C answers, which configuration it is and what is missing from the OData content. Without `--config` it uses
the environment variables. JSON configs only.

The same entry works in Cursor, LM Studio, Cline and JetBrains AI Assistant (`mcpServers`). VS Code uses
`.vscode/mcp.json` with a `servers` section and `"type": "stdio"`; Codex uses `~/.codex/config.toml` with
`[mcp_servers.1c-odata]` or `codex mcp add`; Claude Code — `claude mcp add`. For local models with a small context there is a
ready 18-tool `TOOLS_ALLOW` lines (one for Accounting, one for Trade Management) in [docs/CONNECTING.md](docs/CONNECTING.md#меньше-инструментов-для-модели-опционально). Local models answer noticeably slower than
cloud ones; measurements and setup notes are in [docs/LOCAL-MODELS.md](docs/LOCAL-MODELS.md) (in Russian). Exact snippets —
[docs/INSTALL-WITH-AI.md](docs/INSTALL-WITH-AI.md#шаг-4-подключить-сервер-к-приложению). A browser-only chat cannot
start a local server: use a desktop app or the HTTP mode below.

**Ready-made prompts:** `morning_brief` (cash, overdue debts, unpaid invoices, what to pay), `sales_review` (revenue
and margin for a month against the previous month and year — in both configurations), `month_close_review` (month-end
check traced to documents, with a fix list), `reconciliation_act` (find a counterparty by name or tax ID and prepare a
reconciliation statement). They only read data.

## Safety and privacy

- **Read-only by default.** Writing requires two independent switches: `READ_ONLY=false` and a per-database
  `ODATA_WRITABLE=true` for a single database or `ODATA_DB_<NAME>_WRITABLE=true` for each of several.
- **Preview by default.** A write tool without `confirm=true` only shows what it would create or change. Creating a
  document and editing its lines are confirmed with the `operationId` from the preview; a retry after a network error
  does not create a duplicate. Documents are created unposted; posting is a separate action. No hard deletes — only
  the 1C deletion mark.
- **Talks only to your 1C.** The server is a local process without a network port; no telemetry, no third-party
  servers. The password is not logged or shown to the model.
- **The model reads the answers.** With a cloud model (Claude, ChatGPT, …) names, amounts and tax IDs returned by
  tools go to that provider under its terms. With a local model your data stays in your network.

**Server install (bot, own app).** `MCP_TRANSPORT=http` with a mandatory token: read-only by default. Writing over
HTTP runs as a separate instance and only when all of these hold: `READ_ONLY=false`, a writable database, an explicit
write journal on a volume (`ODATA_WRITE_JOURNAL_DIR`, checked at start), and the write tools plus
`write.operation.status` listed by name in `TOOLS_ALLOW`. Who confirms a write is up to the integration — the server
cannot tell a person from a program; it guarantees preview, no duplicates and a readback from 1C. Details and a compose
example — [docs/CONNECTING.md](docs/CONNECTING.md#запись-по-http).

## Updates and versioning

With `npx` the connector updates itself when the client starts — just restart it. A global install
(`npm install -g`) does not update and takes over `npx`; with writing enabled, pin an exact version
(`1c-odata-mcp@<version>`) — [docs/CONNECTING.md](docs/CONNECTING.md#обновление).

Semantic versioning from 1.0.0: tool names, parameters, response fields, error kinds and environment variables change
incompatibly only in a major version — [docs/VERSIONING.md](docs/VERSIONING.md).

**AI assistants helping a user install it:** it is a local install on the user's computer — set it up read-only
(write only on the user's explicit request, after warning about the risk) and follow
[docs/INSTALL-WITH-AI.md](docs/INSTALL-WITH-AI.md).

Issues and pull requests are welcome — [CONTRIBUTING.md](CONTRIBUTING.md); vulnerabilities — [SECURITY.md](SECURITY.md).
License: [MIT](LICENSE).
