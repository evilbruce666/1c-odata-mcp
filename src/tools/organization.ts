import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { Connection, ServerContext } from "../context.js";
import { ok, fail, guard, databaseField, organizationField } from "./_shared.js";
import { CATALOGS, REGISTERS, resolveEntity } from "../config/mapping.js";
import { resolveOrgOrDefault } from "../odata/orgs.js";
import { resolveNames } from "../odata/accounting.js";
import { KI_KINDS } from "../odata/refdata.js";
import { fetchAll } from "../odata/pagination.js";
import { cmp, odataGuid, buildQuery } from "../odata/query.js";
import type { ODataEntity } from "../types/odata.js";
import { organizationCardResultSchema } from "../schemas/output.js";

/**
 * Карточка организации (реквизиты) — читает Catalog_Организации целиком и
 * дополняет данными из связанных справочников (банк, статьи КИ) и, если
 * опубликован, периодического регистра «Ответственные лица организаций».
 */

const EMPTY_GUID = "00000000-0000-0000-0000-000000000000";
const isRealGuid = (v: unknown): v is string =>
  typeof v === "string" && v.length > 0 && v.replace(/[{}]/g, "").toLowerCase() !== EMPTY_GUID;

const str = (v: unknown): string => (typeof v === "string" ? v : (v ?? "") === "" ? "" : String(v));

/** Дата 0001-01-01 в 1С означает «не заполнено». */
function dateOnly(v: unknown): string | undefined {
  const s = typeof v === "string" ? v : undefined;
  if (!s || s.startsWith("0001-01-01")) return undefined;
  return s.slice(0, 10);
}

function clean<T extends Record<string, unknown>>(obj: T): Partial<T> {
  return Object.fromEntries(Object.entries(obj).filter(([, v]) => v !== undefined)) as Partial<T>;
}

async function getEntityByRef(
  conn: Connection,
  entitySet: string,
  ref: string,
  select?: string[],
): Promise<ODataEntity> {
  const guid = ref.replace(/[{}']/g, "");
  const path = `${entitySet}(guid'${guid}')${buildQuery(select ? { select } : {})}`;
  return conn.client.getEntity(path);
}

interface ResponsiblePersons {
  director?: { fullName: string };
  accountant?: { fullName: string };
  note?: string;
}

/**
 * Читает директора/гл.бухгалтера из периодического регистра «Ответственные лица
 * организаций». Имена полей у этого регистра могут отличаться между конфигурациями
 * (в обеих проверенных живых базах он не опубликован), поэтому структура
 * определяется динамически по $metadata, а не жёстко зашитыми именами.
 */
async function findResponsiblePersons(
  conn: Connection,
  regSet: string,
  orgRef: string,
  available: ReadonlySet<string>,
): Promise<ResponsiblePersons> {
  const meta = await conn.getMetadata();
  const props = meta.entities.get(regSet)?.properties ?? [];
  const names = props.map((p) => p.name);

  const hasPeriod = names.includes("Period");
  const kindField = names.find((n) => n !== "Организация_Key" && !n.endsWith("_Key") && /Вид/i.test(n));
  const personKeyField = names.find(
    (n) => n.endsWith("_Key") && n !== "Организация_Key" && /ФизическоеЛицо|Сотрудник/i.test(n),
  );
  const fioField = personKeyField ? undefined : names.find((n) => !n.endsWith("_Key") && /ФИО/i.test(n));

  if (!kindField || (!personKeyField && !fioField)) {
    return {
      note: "Регистр «Ответственные лица» опубликован, но его структура не распознана автоматически.",
    };
  }

  const select = [
    "Организация_Key",
    kindField,
    ...(personKeyField ? [personKeyField] : []),
    ...(fioField ? [fioField] : []),
    ...(hasPeriod ? ["Period"] : []),
  ];
  const { rows } = await fetchAll(
    conn.client,
    regSet,
    {
      filter: cmp("Организация_Key", "eq", odataGuid(orgRef)),
      select,
      orderby: hasPeriod ? "Period desc" : undefined,
    },
    200,
    200,
  );

  // Самая свежая запись на каждый вид ответственного (Period desc уже отсортировал).
  const latestByKind = new Map<string, ODataEntity>();
  for (const r of rows) {
    const kind = String(r[kindField] ?? "");
    if (kind && !latestByKind.has(kind)) latestByKind.set(kind, r);
  }

  let personNames = new Map<string, string>();
  if (personKeyField) {
    const keys = [...latestByKind.values()].map((r) => String(r[personKeyField] ?? "")).filter(isRealGuid);
    const personSet = resolveEntity(CATALOGS.physicalPersons, available);
    if (personSet && keys.length) personNames = await resolveNames(conn, personSet, keys);
  }

  const fioOf = (r: ODataEntity): string | undefined => {
    if (personKeyField) return personNames.get(String(r[personKeyField] ?? ""));
    if (fioField) return str(r[fioField]) || undefined;
    return undefined;
  };

  let director: { fullName: string } | undefined;
  let accountant: { fullName: string } | undefined;
  for (const [kind, row] of latestByKind) {
    const fullName = fioOf(row);
    if (!fullName) continue;
    if (/Руководител|Директор/i.test(kind)) director = { fullName };
    else if (/Бухгалтер/i.test(kind)) accountant = { fullName };
  }
  return { director, accountant };
}

export function registerOrganizationTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.organization.get_organization_card",
    {
      title: "Карточка организации (реквизиты)",
      description:
        "Полные реквизиты организации (юрлица) из настроек 1С: ИНН/КПП/ОГРН, полное и " +
        "сокращённое наименование, дата регистрации, ОКВЭД, налоговый орган, контактная " +
        "информация (юридический/фактический/почтовый адрес, телефон, email), основной " +
        "банковский счёт, а также директор и главный бухгалтер. Для ИП директором считается " +
        "сам предприниматель (берётся из карточки " +
        "организации); для ООО и т.п. — из регистра «Ответственные лица организаций», " +
        "если он опубликован в OData (если нет — вернётся заметка, что нужно добавить в " +
        "«Состав OData»). Работает для любой организации базы — по названию (organization) " +
        "или автоматически, если она в базе одна.",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
      },
      outputSchema: organizationCardResultSchema,
    },
    ({ database, organization }) =>
      guard("read.organization.get_organization_card", async () => {
        const conn = ctx.db(database);
        const org = await resolveOrgOrDefault(conn, organization);
        const available = await conn.available();
        const orgSet = resolveEntity(CATALOGS.organizations, available);
        if (!orgSet) return fail("Справочник «Организации» не опубликован в OData.");

        const doc = await getEntityByRef(conn, orgSet, org.ref);
        const notes: string[] = [];

        // ОКВЭД: предпочитаем актуальный ОКВЭД2, откат на устаревший ОКВЭД.
        const okvedCode = str(doc["КодОКВЭД2"]) || str(doc["КодОКВЭД"]);
        const okvedName = str(doc["НаименованиеОКВЭД2"]) || str(doc["НаименованиеОКВЭД"]);
        const taxCode = str(doc["КодНалоговогоОргана"]);
        const taxName = str(doc["НаименованиеНалоговогоОргана"]);

        // Адреса/контакты из КонтактнаяИнформация — расшифровываем Вид_Key именами вида.
        const kiRows = (doc["КонтактнаяИнформация"] as Array<Record<string, unknown>>) ?? [];
        const kiKindKeys = new Set(kiRows.map((r) => String(r["Вид_Key"] ?? "")).filter(isRealGuid));
        const kiSet = resolveEntity(KI_KINDS, available);
        const kiNames =
          kiSet && kiKindKeys.size ? await resolveNames(conn, kiSet, kiKindKeys) : new Map<string, string>();
        const contacts = kiRows
          .filter((r) => r["Представление"])
          .map((r) => ({
            kind: kiNames.get(String(r["Вид_Key"])) ?? str(r["Тип"]) ?? "—",
            value: String(r["Представление"]),
          }));

        // Основной банковский счёт (+ банк по БИК, валюта).
        let bankAccount: Record<string, unknown> | undefined;
        const acctRef = str(doc["ОсновнойБанковскийСчет_Key"]);
        if (isRealGuid(acctRef)) {
          const acctSet = resolveEntity(CATALOGS.bankAccounts, available);
          if (acctSet) {
            try {
              const acct = await getEntityByRef(conn, acctSet, acctRef, [
                "НомерСчета",
                "Банк_Key",
                "ВалютаДенежныхСредств_Key",
              ]);
              let bank: { name: string; bik: string } | undefined;
              const bankKey = str(acct["Банк_Key"]);
              if (isRealGuid(bankKey)) {
                const banksSet = resolveEntity(["Catalog_Банки"], available);
                if (banksSet) {
                  const b = await getEntityByRef(conn, banksSet, bankKey, ["Description", "Code"]);
                  bank = { name: str(b["Description"]), bik: str(b["Code"]) };
                }
              }
              let currency: { name: string; code: string } | undefined;
              const curKey = str(acct["ВалютаДенежныхСредств_Key"]);
              if (isRealGuid(curKey)) {
                const curSet = resolveEntity(CATALOGS.currencies, available);
                if (curSet) {
                  const c = await getEntityByRef(conn, curSet, curKey, ["Description", "Code"]);
                  currency = { name: str(c["Description"]), code: str(c["Code"]) };
                }
              }
              bankAccount = clean({ accountNumber: str(acct["НомерСчета"]) || undefined, bank, currency });
            } catch {
              notes.push("Не удалось прочитать основной банковский счёт.");
            }
          }
        }

        // Директор/бухгалтер.
        const legalType = str(doc["ЮридическоеФизическоеЛицо"]);
        let director: { fullName: string } | undefined;
        let accountant: { fullName: string } | undefined;

        if (legalType === "ФизическоеЛицо") {
          const full = [str(doc["ФамилияИП"]), str(doc["ИмяИП"]), str(doc["ОтчествоИП"])]
            .filter(Boolean)
            .join(" ");
          if (full) director = { fullName: full };
        } else {
          const regSet = resolveEntity(REGISTERS.responsiblePersons, available);
          if (!regSet) {
            notes.push(
              "Регистр «Ответственные лица организаций» не опубликован в OData — директор/бухгалтер " +
                "недоступны. Добавьте его в «Состав OData», если нужны эти реквизиты.",
            );
          } else {
            try {
              const found = await findResponsiblePersons(conn, regSet, org.ref, available);
              director = found.director;
              accountant = found.accountant;
              if (found.note) notes.push(found.note);
            } catch (e) {
              notes.push(
                `Не удалось прочитать ответственных лиц (структура регистра не распознана): ${
                  e instanceof Error ? e.message : String(e)
                }`,
              );
            }
          }
        }

        return ok(
          clean({
            database: conn.cfg.name,
            ref: org.ref,
            name: str(doc["Description"]) || undefined,
            fullName: str(doc["НаименованиеПолное"]) || undefined,
            shortName: str(doc["НаименованиеСокращенное"]) || undefined,
            legalType: legalType || undefined,
            inn: str(doc["ИНН"]) || undefined,
            kpp: str(doc["КПП"]) || undefined,
            ogrn: str(doc["ОГРН"]) || undefined,
            registrationDate: dateOnly(doc["ДатаРегистрации"]),
            okved: okvedCode ? { code: okvedCode, name: okvedName || undefined } : undefined,
            taxAuthority:
              taxName || taxCode
                ? clean({ code: taxCode || undefined, name: taxName || undefined })
                : undefined,
            contacts: contacts.length ? contacts : undefined,
            bankAccount,
            director,
            accountant,
            notes: notes.length ? notes : undefined,
          }),
        );
      }),
  );
}
