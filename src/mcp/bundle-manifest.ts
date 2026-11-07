/**
 * Манифест для .mcpb. В репозитории у инструментов ещё схемы и аннотации: на них держится тест согласованности
 * документации с кодом. Формат MCPB допускает в tools только name и description, и Claude Desktop отвергает
 * расширение с другими ключами («Invalid manifest: tools: Unrecognized key(s)»). Схемы клиент всё равно получает
 * от самого сервера через tools/list.
 */
export function bundleManifest<T extends { tools?: Array<{ name: string; description?: string }> }>(
  manifest: T,
): T {
  return {
    ...manifest,
    tools: manifest.tools?.map((t) => ({ name: t.name, description: t.description })),
  };
}
