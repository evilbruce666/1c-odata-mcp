/**
 * Двоичные поля документа («…_Base64Data»: содержимое отчёта, хранилище значений, вложенные данные) модели не
 * прочитать, а контекст они занимают целиком — регламентированный отчёт весил около 70 тысяч знаков. Убираются только
 * из ответа чтения: запись (copy_document) читает сущность сама, целиком.
 */
const BINARY_SUFFIX = "_Base64Data";

/** Копия без непустых двоичных полей и их пути («Строки[].Файл_Base64Data»), без повторов. */
export function withoutBinary(entity: Record<string, unknown>): {
  entity: Record<string, unknown>;
  omitted: string[];
} {
  const omitted = new Set<string>();
  const walk = (value: unknown, path: string): unknown => {
    if (Array.isArray(value)) return value.map((v) => walk(v, `${path}[]`));
    if (!value || typeof value !== "object") return value;
    return Object.fromEntries(
      Object.entries(value).flatMap(([key, v]) => {
        const at = path ? `${path}.${key}` : key;
        if (key.endsWith(BINARY_SUFFIX) && typeof v === "string" && v !== "") {
          omitted.add(at);
          return [];
        }
        return [[key, walk(v, at)]];
      }),
    );
  };
  return { entity: walk(entity, "") as Record<string, unknown>, omitted: [...omitted] };
}
