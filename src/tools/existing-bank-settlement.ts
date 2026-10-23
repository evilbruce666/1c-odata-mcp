import type { McpServer } from "@modelcontextprotocol/server";
import type { ServerContext } from "../context.js";
import { z } from "zod";
import { applyBankSettlementSettings, bankSettlementInput } from "../odata/existing-bank-settlement.js";
import { databaseField, guard, ok } from "./_shared.js";
export function registerExistingBankSettlementTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "write.document.apply_bank_settlement_settings",
    {
      title: "Применить счета расчётов к существующему банковскому документу",
      description:
        "Применяет унаследованные счета расчётов/авансов и выбранный договор к существующему НЕПРОВЕДЁННОМУ банковскому документу с точной expectedDataVersion/If-Match. Явная сторона supplier/customer обязательна, включая договор Прочее. Необязательный counterpartyRef заполняет только нулевую ссылку контрагента документа; требует существующий неудалённый элемент и точного владельца договора. Замена заполненного контрагента и подбор по названию/договору запрещены. Поддерживает опубликованные GUID и составной тип контрагента. Контрольное чтение проверяет договор, оба счёта, суммы платежа, налог и аналитику; пересчитанные курс, кратность и сумма взаиморасчётов не сравниваются с прежними. Условные единицы с иной валютой документа пока не поддержаны, курс не угадывается; несколько строк отклоняются. Предпросмотр связан с payload, документом, явно выбранным контрагентом, договором и результатом подбора. Отмена проведения и проведение выполняются отдельными вызовами; unknown удерживается без повторной записи.",
      inputSchema: z.object({
        database: databaseField,
        ...bankSettlementInput,
        confirm: z.boolean().default(false),
      }),
      outputSchema: z.object({}).passthrough(),
    },
    (args) =>
      guard("write.document.apply_bank_settlement_settings", async () =>
        ok(await applyBankSettlementSettings(ctx.db(args.database), args, args.confirm)),
      ),
  );
}
