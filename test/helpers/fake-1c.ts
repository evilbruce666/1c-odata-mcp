import { vi } from "vitest";
import { parseConfig } from "../../src/config/env.js";
import { ServerContext, type Connection } from "../../src/context.js";
import { BP30_METADATA_XML, bp30Enums, bp30Metadata, SNAPSHOT_REF } from "./metadata-snapshot.js";

/**
 * «Поддельная 1С» на снимке схемы БП 3.0: подменяет fetch и отвечает как OData 1С. Коннектор
 * работает целиком (загрузка $metadata, клиент, классификация ошибок), а подделка проверяет каждый
 * запрос так же строго, как настоящая база: объект существует, поля $select / $filter / $orderby
 * есть, в теле записи (и в строках табличных частей) только опубликованные поля, функция регистра и
 * её параметры объявлены в $metadata, вид операции — значение перечисления. Нарушение — ответ 400 с текстом, как у 1С, и запись в `violations`.
 * Данные синтетические: по одной строке каждого объекта со значениями по типам полей.
 */

export const FAKE_BASE_URL = "https://fake-1c.invalid/odata/standard.odata/";

export interface FakeRequest {
  method: string;
  resource: string;
}

export interface Fake1C {
  ctx: ServerContext;
  conn: Connection;
  requests: FakeRequest[];
  violations: string[];
}

/** Функции регистров из снимка (Balance, Turnovers, SliceLast…): «регистр/имя» → параметры и тип результата. */
interface FakeFunction {
  params: Set<string>;
  returns: string;
}
let functionCache: Map<string, FakeFunction> | undefined;
function functions(): Map<string, FakeFunction> {
  if (functionCache) return functionCache;
  functionCache = new Map();
  for (const m of BP30_METADATA_XML.matchAll(/<FunctionImport Name="([^"]+)"[\s\S]*?<\/FunctionImport>/g)) {
    const params = [...m[0].matchAll(/<Parameter Name="([^"]+)"\s+Type="([^"]+)"/g)];
    const binding = params.find((p) => p[1] === "bindingParameter")?.[2]?.replace(/^StandardODATA\./, "");
    const returns = /ReturnType="(?:Collection\()?StandardODATA\.([^")]+)/.exec(m[0])?.[1];
    if (binding && returns)
      functionCache.set(`${binding}/${m[1]}`, {
        params: new Set(params.map((p) => p[1]!).filter((n) => n !== "bindingParameter")),
        returns,
      });
  }
  return functionCache;
}

/** Слова языка запросов OData, которые не являются полями. */
const KEYWORDS = new Set(
  (
    "eq ne gt ge lt le and or not true false null guid datetime datetimeoffset time binary " +
    "substringof startswith endswith cast isof year month day hour minute second tolower toupper " +
    "length indexof trim concat replace substring round floor ceiling add sub mul div mod asc desc " +
    "X M L Edm StandardODATA"
  ).split(" "),
);

let complexCache: Map<string, Map<string, string>> | undefined;
function complexTypes(): Map<string, Map<string, string>> {
  if (complexCache) return complexCache;
  complexCache = new Map();
  for (const m of BP30_METADATA_XML.matchAll(/<ComplexType Name="([^"]+)"[\s\S]*?<\/ComplexType>/g))
    complexCache.set(
      m[1]!,
      new Map([...m[0].matchAll(/<Property Name="([^"]+)"\s+Type="([^"]+)"/g)].map((p) => [p[1]!, p[2]!])),
    );
  return complexCache;
}

function entityFields(set: string): Map<string, string> | undefined {
  const meta = bp30Metadata().entities.get(set);
  if (!meta) return undefined;
  return new Map([
    ...meta.properties.map((p) => [p.name, p.type] as const),
    ...meta.navigations.map((n) => [n.name, "navigation"] as const),
  ]);
}

const operationKinds = (): Set<string> =>
  new Set([...bp30Enums()].filter(([n]) => n.startsWith("ВидыОпераций")).flatMap(([, v]) => [...v]));

/** Синтетическое значение по типу Edm (ключи — тот же GUID, чтобы проверки «тот ли объект» проходили). */
const EMPTY_GUID = "00000000-0000-0000-0000-000000000000";

function syntheticValue(name: string, type: string): unknown {
  if (type.startsWith("Collection(")) return [];
  if (name === "Parent_Key") return EMPTY_GUID; // объекты в корне: иерархия без циклов
  switch (type) {
    case "Edm.Guid":
      return SNAPSHOT_REF;
    case "Edm.Boolean":
      return false;
    case "Edm.DateTime":
      return "2026-09-01T00:00:00";
    case "Edm.Double":
    case "Edm.Decimal":
    case "Edm.Int16":
    case "Edm.Int32":
    case "Edm.Int64":
      return 0;
    default:
      return OVERRIDES[name] ?? (name === "Description" ? "Synthetic" : "");
  }
}

const OVERRIDES: Record<string, unknown> = {
  Контрагент: SNAPSHOT_REF,
  Контрагент_Type: "StandardODATA.Catalog_Контрагенты",
  Owner_Type: "StandardODATA.Catalog_Контрагенты",
  ВидДоговора: "СПокупателем",
  Code: "643",
  Number: "0000-000001",
  DataVersion: "AAAAAAAAAAA=",
  ВидОперации: "ОплатаПокупателя",
};

/** Коды плана счетов, которые ищут инструменты (по одному синтетическому счёту на код). */
const ACCOUNT_CODES = [
  "41.01",
  "43",
  "10.01",
  "50.01",
  "51",
  "52",
  "55",
  "57.03",
  "60.01",
  "60.02",
  "62.01",
  "62.02",
  "76.09",
  "90.01.1",
  "90.02.1",
  "90.03",
  "90.09",
  "91.01",
  "91.02",
  "94",
  "26",
  "44.01",
  "68.02",
  "19.03",
];

const GOODS_SECTIONS = new Set(["Товары", "Услуги"]);

function syntheticRow(fields: Map<string, string>, set: string): Record<string, unknown> {
  const row: Record<string, unknown> = {};
  for (const [name, type] of fields) {
    if (type === "navigation") continue;
    if (type.startsWith("Collection(")) {
      const section = entityFields(`${set}_${name}`);
      const r = section ? syntheticRow(section, `${set}_${name}`) : undefined;
      // Товары и услуги — две строки (есть что удалить), прочие табличные части — одна.
      const lines = GOODS_SECTIONS.has(name) ? ["1", "2"] : ["1"];
      row[name] = r ? lines.map((LineNumber) => ({ ...r, LineNumber })) : [];
    } else row[name] = syntheticValue(name, type);
  }
  return row;
}

function json(status: number, body: unknown): Response {
  return new Response(body === undefined ? "" : JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

const odataError = (status: number, text: string) =>
  json(status, { "odata.error": { code: "-1", message: { lang: "ru", value: text } } });

/** Имена полей в выражении $filter / $orderby (без строк, литералов и ключевых слов). */
function identifiers(expr: string): string[] {
  const noStrings = expr.replace(/'(?:[^']|'')*'/g, " ");
  return [...noStrings.matchAll(/[A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё]*(?:\/[\wА-Яа-яЁё]+)*/g)]
    .map((m) => m[0]!)
    .filter((w) => !KEYWORDS.has(w) && !/^\d/.test(w));
}

export function fake1C(env: Record<string, string> = {}): Fake1C {
  const requests: FakeRequest[] = [];
  const violations: string[] = [];
  const kinds = operationKinds();
  /** Признак и версия проведённых в этом прогоне документов, по пути «Набор(guid'…')». */
  const postedState = new Map<string, { Posted: boolean; DataVersion: string }>();
  /** Поля шапки, изменённые PATCH в этом прогоне, по тому же пути. */
  const storedState = new Map<string, Record<string, unknown>>();

  const handle = (url: string, method: string, body: unknown): Response => {
    const raw = url.startsWith(FAKE_BASE_URL) ? url.slice(FAKE_BASE_URL.length) : url;
    const [pathPart = "", query = ""] = raw.split("?");
    const path = decodeURIComponent(pathPart);
    const params = new URLSearchParams(query);
    if (path === "$metadata")
      return new Response(BP30_METADATA_XML, { status: 200, headers: { "Content-Type": "application/xml" } });

    const base = /^[^(/]+/.exec(path)?.[0] ?? path;
    requests.push({ method, resource: base });
    if (process.env.FAKE_1C_TRACE)
      process.stderr.write(`[fake-1c] ${method} ${path}?${decodeURIComponent(query)}\n`);
    const fields = entityFields(base);
    if (!fields) return odataError(404, `Сущность '${base}' не найдена`);

    // Виртуальная таблица регистра: поля — из её типа; действие документа — без полей.
    const tail = path.slice(base.length).replace(/^\([^)]*\)/, "");
    const call = /^\/([\wА-Яа-яЁё]+)(?:\(([^)]*)\))?/.exec(tail);
    const fn = call?.[1];
    const declared = fn ? functions().get(`${base}/${fn}`) : undefined;
    const bad = (where: string, name: string) => {
      violations.push(`${method} ${base}${fn ? `/${fn}` : ""}: ${where} — нет поля «${name}»`);
      return odataError(400, `Неизвестное поле '${name}' (${where})`);
    };
    if (fn && !declared && !["Post", "Unpost"].includes(fn) && !fields.has(fn)) {
      violations.push(`${method} ${base}/${fn}: функция не объявлена в $metadata`);
      return odataError(400, `Неизвестное действие или поле '${fn}' у '${base}'`);
    }
    // Параметры функции — только объявленные (Period, StartPeriod, Condition…).
    if (declared)
      for (const p of (call?.[2] ?? "").replace(/'(?:[^']|'')*'/g, "''").matchAll(/([\wА-Яа-яЁё]+)=/g))
        if (!declared.params.has(p[1]!)) return bad(`параметр ${fn}()`, p[1]!);
    const virtual = declared
      ? (complexTypes().get(declared.returns) ?? entityFields(declared.returns))
      : undefined;
    const scope = virtual ?? fields;
    const known = (name: string) => scope.has(name.split("/")[0]!);

    for (const name of (params.get("$select") ?? "").split(",").filter(Boolean))
      if (!known(name.trim())) return bad("$select", name.trim());
    for (const key of ["$filter", "$orderby"] as const)
      for (const name of identifiers(params.get(key) ?? "")) if (!known(name)) return bad(key, name);

    if (method === "POST" && (fn === "Post" || fn === "Unpost")) {
      // Как в 1С: проведение меняет признак и версию документа — инструмент сверяет их после действия.
      const doc = /^[^(]+\([^)]*\)/.exec(path)![0];
      const version = Number(postedState.get(doc)?.DataVersion ?? 0) + 1;
      postedState.set(doc, { Posted: fn === "Post", DataVersion: String(version) });
      return json(200, undefined);
    }
    if (method === "POST" || method === "PATCH") {
      const payload = (body ?? {}) as Record<string, unknown>;
      for (const [key, value] of Object.entries(payload)) {
        if (key.includes("@")) continue;
        if (!fields.has(key)) return bad("тело записи", key);
        if (Array.isArray(value)) {
          const section = entityFields(`${base}_${key}`);
          for (const row of value as Record<string, unknown>[])
            for (const k of Object.keys(row ?? {})) if (!section?.has(k)) return bad(`строка ${key}`, k);
        }
        if (key === "ВидОперации" && typeof value === "string" && !kinds.has(value)) {
          violations.push(`${method} ${base}: вид операции «${value}» не входит ни в одно перечисление`);
          return odataError(500, `Перечисление 'ВидыОпераций' не содержит элемент '${value}'`);
        }
      }
      // Изменение запоминается: повторное чтение объекта видит сохранённые поля, как в 1С.
      if (method === "PATCH") {
        const doc = /^[^(]+\([^)]*\)/.exec(path)?.[0];
        if (doc) {
          const scalars = Object.fromEntries(Object.entries(payload).filter(([, v]) => !Array.isArray(v)));
          storedState.set(doc, { ...storedState.get(doc), ...scalars });
        }
      }
      return json(method === "POST" ? 201 : 200, { ...syntheticRow(fields, base), ...payload });
    }

    if (virtual) return json(200, { value: [syntheticRow(virtual, base)] });
    const keyed = tail === "" && /^[^(]+\(/.test(path);
    if (Number(params.get("$skip") ?? 0) > 0 && !keyed) return json(200, { value: [] });
    if (base.startsWith("ChartOfAccounts_") && !keyed)
      return json(200, {
        value: ACCOUNT_CODES.map((Code) => ({ ...syntheticRow(fields, base), Code, Description: Code })),
      });
    // Постраничное чтение: синтетика — одна страница, дальше пусто (иначе $skip крутился бы вечно).
    if (Number(params.get("$skip") ?? 0) > 0) return json(200, { value: [] });
    const row = syntheticRow(fields, base);
    const guid = /guid'([0-9a-f-]{36})'/i.exec(path)?.[1];
    if (keyed && guid && "Ref_Key" in row) row.Ref_Key = guid;
    if (keyed) Object.assign(row, storedState.get(path), postedState.get(path));
    // Составной ключ (регистры сведений): запись отвечает теми значениями измерений, по которым её запросили.
    const key = /^[^(]+\(([^)]*)\)/.exec(path)?.[1] ?? "";
    for (const m of key.matchAll(/([\wА-Яа-яЁё]+)=(?:guid'([0-9a-f-]{36})'|'((?:[^']|'')*)')/gi))
      if (m[1]! in row) row[m[1]!] = m[2] ?? m[3]!.replace(/''/g, "'");
    // Отбор «поле eq значение»: строка коллекции ему удовлетворяет, как в настоящей базе.
    if (!keyed)
      for (const m of (params.get("$filter") ?? "").matchAll(
        /([\wА-Яа-яЁё]+) eq (?:guid'([0-9a-f-]{36})'|'((?:[^']|'')*)'|cast\(guid'([0-9a-f-]{36})',\s*'([^']+)'\))/gi,
      )) {
        if (!(m[1]! in row)) continue;
        row[m[1]!] = m[2] ?? m[4] ?? m[3]!.replace(/''/g, "'");
        if (m[5] && `${m[1]}_Type` in row) row[`${m[1]}_Type`] = `StandardODATA.${m[5]}`;
      }
    return json(200, keyed ? row : { value: [row] });
  };

  vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
    const method = (init?.method ?? "GET").toUpperCase();
    const body = typeof init?.body === "string" ? (JSON.parse(init.body) as unknown) : undefined;
    return handle(url, method, body);
  });

  const ctx = new ServerContext(
    parseConfig({
      ODATA_BASE_URL: FAKE_BASE_URL,
      ODATA_USERNAME: "u",
      ODATA_PASSWORD: "p",
      ODATA_RETRIES: "0",
      ...env,
    }),
  );
  return { ctx, conn: ctx.db(), requests, violations };
}
