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

1. Пересобрать список инструментов в `manifest.json` из живого сервера (имя, описание,
   `inputSchema`, `outputSchema`, аннотации — ровно то, что видит MCP-клиент) и
   синхронизировать `version` с `package.json`. Реальная база не нужна — сервер
   поднимается с фиктивным подключением в режиме только-чтения:

   ```bash
   npm run manifest
   ```

2. Прогнать тесты: `test/docs-consistency.test.ts` сверяет состав и схемы инструментов в
   манифесте с кодом, версию в `package.json`/`manifest.json`/CHANGELOG и число
   инструментов в README, `package.json` и `long_description` манифеста. Упал — поправить
   числа в текстах (`mcpb validate` тут не поможет: официальный валидатор не принимает
   `inputSchema` в `tools[]`):

   ```bash
   npm test
   ```

3. Собрать staging-папку (build + **чистый** `npm ci --omit=dev`, чтобы не
   тащить dev-зависимости и не трогать рабочий `node_modules`) и заzip'овать:

   ```bash
   STAGE=$(mktemp -d)
   npm run build
   cp -R dist "$STAGE/dist"
   cp package.json package-lock.json README.md LICENSE manifest.json "$STAGE/"
   (cd "$STAGE" && npm ci --omit=dev --ignore-scripts)
   rm -rf "$STAGE/dist/scripts"
   # В рантайме не нужны: типы, sourcemaps, исходники .ts и markdown пакетов —
   # бандл ужимается с ~6,4 до ~3,3 МБ. Иначе на медленном канале (~100 КБ/с)
   # `smithery mcp publish` не успевает загрузить файл и падает «Request timed out».
   find "$STAGE" -type f \( -name "*.map" -o -name "*.d.ts" -o -name "*.d.mts" -o -name "*.d.cts" \) -delete
   find "$STAGE/node_modules" -type f \( -name "*.ts" -o -iname "*.md" -o -iname "CHANGELOG*" \) -delete
   (cd "$STAGE" && zip -r -X -q -9 /tmp/1c-odata-mcp.mcpb . -x ".*")
   ```

   Перед публикацией стоит запустить `$STAGE/dist/index.js` MCP-клиентом (`tools/list` и
   пара вызовов) — облегчённый бандл должен отвечать так же, как исходный.

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
