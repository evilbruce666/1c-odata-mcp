/**
 * Сравнение номеров документов 1С.
 *
 * Свой номер документа — префикс и номер с ведущими нулями разной длины: «0000-000033»,
 * «00000000001», «0000-00000024», «БП-000033». Человек называет его без них — «№33», поэтому
 * для запроса из одних цифр совпадение — по последней группе цифр без нулей. Нумерация в 1С
 * начинается заново каждый год и у каждой организации своя: «№33» — обычно несколько документов.
 *
 * Входящий номер (поставщика, банка контрагента) бывает любым: «ГК-22537», «20290930-29»,
 * «0000-05893», с пробелом в конце. Его последняя группа цифр — не номер («29» у «20290930-29»),
 * поэтому совпадение — только целиком (без пробелов и регистра) или, если оба из одних цифр,
 * без ведущих нулей.
 */

export interface NumberQuery {
  /** Запрос без «№», пробелов и регистра — для сравнения. */
  normalized: string;
  /** Запрос из одних цифр без ведущих нулей; undefined — в запросе есть не только цифры. */
  digits?: string;
  /** Подстрока для отбора на сервере 1С (substringof/endswith регистр не учитывают). */
  needle: string;
}

const normalize = (s: string) => s.replace(/\s+/g, "").toLowerCase();
const stripZeros = (d: string) => d.replace(/^0+(?=\d)/, "");

export function parseNumberQuery(input: string): NumberQuery | undefined {
  const trimmed = input
    .trim()
    .replace(/^(№|#)\s*/, "")
    .trim();
  const normalized = normalize(trimmed);
  if (!normalized) return undefined;
  const digits = /^\d+$/.test(normalized) ? stripZeros(normalized) : undefined;
  return { normalized, ...(digits !== undefined ? { digits } : {}), needle: digits ?? trimmed };
}

/** Свой номер документа (поле Number). */
export function matchesOwnNumber(value: unknown, q: NumberQuery): boolean {
  const v = normalize(String(value ?? ""));
  if (!v) return false;
  if (v === q.normalized) return true;
  if (q.digits === undefined) return false;
  const tail = /(\d+)$/.exec(v)?.[1];
  return tail !== undefined && stripZeros(tail) === q.digits;
}

/** Входящий номер (номер документа контрагента). */
export function matchesIncomingNumber(value: unknown, q: NumberQuery): boolean {
  const v = normalize(String(value ?? ""));
  if (!v) return false;
  if (v === q.normalized) return true;
  return q.digits !== undefined && /^\d+$/.test(v) && stripZeros(v) === q.digits;
}
