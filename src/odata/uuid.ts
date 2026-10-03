/**
 * Разбор GUID (Ref_Key 1С) по RFC 4122 — только то, что нужно для аудита.
 *
 * UUID версии 1 содержит 60-битную метку времени (интервалы по 100 нс от
 * 1582-10-15T00:00:00Z) и поле node. Если Ref_Key — UUIDv1 варианта RFC 4122,
 * из него можно ВЫВЕСТИ момент генерации ссылки. Это НЕ проверенное время
 * создания/выполнения/проведения документа:
 *  - ссылка могла быть задана программно или перенесена из другой базы (обмен,
 *    загрузка) — тогда это время генерации в источнике;
 *  - корректность часов и часовой пояс генерирующей машины через OData не проверить;
 *  - повторная запись/проведение документа ссылку не меняет.
 * Любая другая версия/вариант — метки времени нет, ничего не выводим.
 */

const UUID_RE = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;

/** 100-нс интервалов между 1582-10-15T00:00:00Z и 1970-01-01T00:00:00Z. */
const GREGORIAN_TO_UNIX_100NS = 122_192_928_000_000_000n;

export type UuidVariant = "NCS" | "RFC4122" | "Microsoft" | "Reserved";

export interface UuidV1Info {
  kind: "v1";
  version: 1;
  variant: "RFC4122";
  /** Метка времени, UTC, ISO-8601 с миллисекундами (дробь < 1 мс отбрасывается). */
  timestampIso: string;
  /** Исходная 60-битная метка (100 нс от 1582-10-15Z) — без потери точности. */
  timestamp100ns: string;
  clockSequence: number;
  /** Поле node (12 hex). Не доказывает идентичность физического сервера. */
  node: string;
  /** Бит multicast в node: по RFC 4122 означает случайный node, а не MAC-адрес. */
  nodeMulticastBit: boolean;
}

export interface UuidNoTimestamp {
  kind: "no-timestamp";
  version: number;
  variant: UuidVariant;
  reason: string;
}

export type UuidTimestampResult = UuidV1Info | UuidNoTimestamp;

function variantOf(byte8: number): UuidVariant {
  if ((byte8 & 0x80) === 0) return "NCS";
  if ((byte8 & 0xc0) === 0x80) return "RFC4122";
  if ((byte8 & 0xe0) === 0xc0) return "Microsoft";
  return "Reserved";
}

/**
 * Разбирает GUID и, только для UUIDv1 варианта RFC 4122, возвращает метку времени.
 * Некорректная строка — исключение (не «нет метки»).
 */
export function uuidTimestamp(value: string): UuidTimestampResult {
  const v = value.trim().replace(/^\{|\}$/g, "");
  if (!UUID_RE.test(v)) throw new Error(`Некорректный GUID: ${value}`);
  const hex = v.replace(/-/g, "").toLowerCase();
  const version = parseInt(hex[12]!, 16);
  const byte8 = parseInt(hex.slice(16, 18), 16);
  const variant = variantOf(byte8);

  if (variant !== "RFC4122") {
    return {
      kind: "no-timestamp",
      version,
      variant,
      reason: `Ref_Key не является UUID варианта RFC 4122 (вариант ${variant}) — метку времени не извлекаем.`,
    };
  }
  if (version !== 1) {
    return {
      kind: "no-timestamp",
      version,
      variant,
      reason: `Ref_Key — UUID версии ${version}, а не версии 1: метки времени в нём нет.`,
    };
  }

  const timeLow = hex.slice(0, 8);
  const timeMid = hex.slice(8, 12);
  const timeHi = hex.slice(13, 16); // без полубайта версии
  const ticks = BigInt(`0x${timeHi}${timeMid}${timeLow}`);
  const unixMs = (ticks - GREGORIAN_TO_UNIX_100NS) / 10_000n;
  const node = hex.slice(20, 32);

  return {
    kind: "v1",
    version: 1,
    variant: "RFC4122",
    timestampIso: new Date(Number(unixMs)).toISOString(),
    timestamp100ns: ticks.toString(),
    clockSequence: parseInt(hex.slice(16, 20), 16) & 0x3fff,
    node,
    nodeMulticastBit: (parseInt(node.slice(0, 2), 16) & 0x01) === 1,
  };
}
