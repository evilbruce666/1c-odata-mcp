import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { InputError } from "../errors.js";
import { resolveOrgOrDefault } from "../odata/orgs.js";
import { buildTaxCalendar } from "../odata/tax-calendar.js";
import { TAX_REGIME_KEYS } from "../odata/tax-regime.js";
import { taxCalendarResultSchema } from "../schemas/output.js";
import { databaseField, guard, ok, organizationField } from "./_shared.js";
import { localDate } from "./local-date.js";

/** Окно по умолчанию: ближайший квартал сроков. */
const DEFAULT_DAYS = 90;
const MAX_DAYS = 731;

const day = z
  .string()
  .regex(/^\d{4}-\d{2}-\d{2}$/, "Дата в формате YYYY-MM-DD")
  .refine((s) => !Number.isNaN(Date.parse(s)) && new Date(s).toISOString().startsWith(s), "Такой даты нет");

export function registerTaxCalendarTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.accounting.get_tax_calendar",
    {
      title: "Налоговый календарь: что заплатить и сдать",
      description:
        "Что организации или ИП заплатить и сдать в окне дат и что из этого уже есть в 1С («какие налоги и " +
        "когда платить», «когда сдавать декларацию по УСН», «сходятся ли уведомления с начислениями»). Только " +
        "Бухгалтерия. Суммы и сроки начисленного (status=accrued) — из 1С; чего в 1С ещё нет (декларации, " +
        "уведомления, будущие взносы и авансы), идёт из справочника со статьёй НК РФ и без сумм — не считайте и " +
        "не придумывайте их сами. Оплачено ли — календарь не знает. prepared — отчёт или уведомление есть в 1С, " +
        "это не значит «сдан», но и не «не сдан». unknown — данных нет, это не «не найдено». Обязательно " +
        "передайте пользователю warnings и missing; справочник покрывает не всё — перечень в notCovered. " +
        "vatThreshold — доход по книгам доходов 1С и порог НДС на УСН: не складывайте и не оценивайте право на " +
        "освобождение сами; нет exceeded — доход учтён не весь. " +
        "По умолчанию — от сегодня на 90 дней.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        from: day.optional().describe("Первый срок окна (YYYY-MM-DD), по умолчанию сегодня"),
        to: day
          .optional()
          .describe("Последний срок окна (YYYY-MM-DD), по умолчанию через 90 дней после from"),
        regime: z
          .array(z.enum(TAX_REGIME_KEYS))
          .min(1)
          .optional()
          .describe(
            "Режим со слов пользователя, если 1С его не отдаёт или отдаёт неверно: usn_income (УСН «доходы»), " +
              "usn_income_expense («доходы минус расходы»), psn (патент), osno. Заменяет определение по базе",
          ),
      }),
      outputSchema: taxCalendarResultSchema,
    },
    ({ database, organization, from, to, regime }) =>
      guard("read.accounting.get_tax_calendar", async () => {
        const today = localDate(new Date());
        const start = from ?? today;
        const end = to ?? new Date(Date.parse(start) + DEFAULT_DAYS * 86_400_000).toISOString().slice(0, 10);
        const days = (Date.parse(end) - Date.parse(start)) / 86_400_000;
        if (days < 0) throw new InputError("to раньше from.");
        if (days > MAX_DAYS) throw new InputError("Окно календаря — не больше двух лет.");
        const conn = ctx.db(database);
        const org = await resolveOrgOrDefault(conn, organization);
        const calendar = await buildTaxCalendar(conn, {
          orgRef: org.ref,
          today,
          from: start,
          to: end,
          regime,
        });
        return ok({ database: conn.cfg.name, organization: { ref: org.ref, name: org.name }, ...calendar });
      }),
  );
}
