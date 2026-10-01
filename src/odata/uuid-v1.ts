import { randomBytes } from "node:crypto";

/** 100-нс интервалов между 1582-10-15T00:00:00Z (эпоха UUID) и 1970-01-01T00:00:00Z. */
const GREGORIAN_TO_UNIX_100NS = 122_192_928_000_000_000n;

let lastTimestamp = 0n;

/**
 * UUID версии 1 (RFC 4122): метка времени + случайные clock sequence и node с битом multicast
 * (§4.5 — признак случайного node, не MAC-адреса). Им MCP задаёт Ref_Key создаваемых объектов:
 * 1С сама выдаёт ссылки v1, поэтому и наши остаются «датируемыми» (время создания выводится
 * из ссылки так же, как у объектов, созданных в 1С), а повтор с той же ссылкой 1С отвергает.
 */
export function uuidV1(now: number = Date.now()): string {
  let ts = BigInt(now) * 10_000n + GREGORIAN_TO_UNIX_100NS;
  // Монотонность внутри процесса: два вызова в одну миллисекунду — разные метки.
  if (ts <= lastTimestamp) ts = lastTimestamp + 1n;
  lastTimestamp = ts;

  const rnd = randomBytes(8);
  const hex = (v: bigint, width: number) => v.toString(16).padStart(width, "0");
  const timeLow = ts & 0xffff_ffffn;
  const timeMid = (ts >> 32n) & 0xffffn;
  const timeHi = ((ts >> 48n) & 0x0fffn) | 0x1000n; // версия 1
  const clockSeq = ((rnd.readUInt16BE(0) & 0x3fff) | 0x8000).toString(16); // вариант RFC 4122
  rnd[2] = rnd[2]! | 0x01; // multicast-бит: node случайный
  const node = rnd.subarray(2, 8).toString("hex");
  return `${hex(timeLow, 8)}-${hex(timeMid, 4)}-${hex(timeHi, 4)}-${clockSeq}-${node}`;
}
