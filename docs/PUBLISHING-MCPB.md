# Обновление листинга на Smithery (MCPB-бандл)

Сервер опубликован на [smithery.ai/servers/alexei/1c-odata-mcp](https://smithery.ai/servers/alexei/1c-odata-mcp)
как **Local**-сервер (MCPB-бандл, https://github.com/modelcontextprotocol/mcpb) —
Smithery раздаёт `.mcpb`, пользователи запускают его у себя, реальные креды
к 1С никогда не попадают на инфраструктуру Smithery.

Источник правды — `manifest.json` в корне репозитория. При выпуске новой версии
(новые инструменты, изменённые описания) его нужно обновить и переопубликовать
вручную — Smithery не делает этого автоматически при пуше в GitHub.

## Почему `mcpb pack` нельзя использовать напрямую

Официальный CLI `@anthropic-ai/mcpb` валидирует `manifest.json` по спецификации
MCPB, где элементы массива `tools` — это только `{name, description}` (без
`inputSchema`). Но бэкенд Smithery, чтобы отрисовать список возможностей
(«Capabilities») на карточке сервера, ожидает у каждого элемента `tools[]`
ещё и `inputSchema` (как в MCP-протоколе) — без него `smithery mcp publish`
отвечает `400 Invalid input: expected object, received undefined` по числу
инструментов. `mcpb pack` при виде `inputSchema` в `tools[]` отказывается
паковать («Unrecognized key(s)»). Поэтому бандл собирается вручную через `zip`,
в обход валидатора CLI — сам файл `.mcpb` это просто zip-архив.

## Как обновить

1. Собрать проект и вытащить актуальный список инструментов из живого сервера
   (имя, описание, `inputSchema`) — самый надёжный источник, совпадает с тем,
   что видит MCP-клиент:

   ```bash
   npm run build
   node --input-type=module -e '
   import("@modelcontextprotocol/sdk/client/index.js").then(async ({Client}) => {
     const {StdioClientTransport} = await import("@modelcontextprotocol/sdk/client/stdio.js");
     const env = {...process.env, ODATA_BASE_URL:"https://example.com/odata/standard.odata/",
       ODATA_USERNAME:"user", ODATA_PASSWORD:"pass", READ_ONLY:"true"};
     const transport = new StdioClientTransport({command:"node", args:["dist/index.js"], env});
     const client = new Client({name:"tool-lister", version:"0.1.0"});
     await client.connect(transport);
     const {tools} = await client.listTools();
     console.log(JSON.stringify(tools.map(t=>({name:t.name, description:t.description, inputSchema:t.inputSchema})), null, 1));
     await client.close();
   })'
   ```

2. Обновить поле `tools` в `manifest.json` этим списком (и `version`, если
   меняли — держите в синхроне с `package.json`). Проверить схему (без `tools`,
   т.к. официальный валидатор его не примет с `inputSchema` — проверяйте
   вручную/JSON.parse, `mcpb validate` тут не поможет):

   ```bash
   node -e "JSON.parse(require('fs').readFileSync('manifest.json','utf8')); console.log('ok')"
   ```

3. Собрать staging-папку (build + **чистый** `npm ci --omit=dev`, чтобы не
   тащить dev-зависимости и не трогать рабочий `node_modules`) и заzip'овать:

   ```bash
   STAGE=$(mktemp -d)
   npm run build
   cp -R dist "$STAGE/dist"
   cp package.json package-lock.json README.md LICENSE manifest.json "$STAGE/"
   (cd "$STAGE" && npm ci --omit=dev --ignore-scripts)
   (cd "$STAGE" && zip -r -X -q /tmp/1c-odata-mcp.mcpb . -x ".*")
   ```

4. Опубликовать (нужен вход `npx @smithery/cli auth login`):

   ```bash
   npx @smithery/cli mcp publish /tmp/1c-odata-mcp.mcpb -n alexei/1c-odata-mcp
   ```

5. Если менялось поле `description`/`homepage`/`license` верхнего уровня — их
   надо обновить отдельным вызовом (Smithery не читает `description` из
   `manifest.json` для Local-бандлов, только `configSchema` и `tools`):

   ```bash
   curl -X PATCH "https://api.smithery.ai/servers/alexei%2F1c-odata-mcp" \
     -H "Authorization: Bearer $(npx @smithery/cli auth token --policy '{"resources":"servers","operations":"write","namespaces":"alexei","ttl":"10m"}' | node -pe "JSON.parse(require('fs').readFileSync(0,'utf8')).token")" \
     -H "Content-Type: application/json" \
     -d '{"description": "..."}'
   ```

6. Проверить карточку: https://smithery.ai/servers/alexei/1c-odata-mcp —
   должны быть описание и все инструменты (не «No capabilities found»).
