/**
 * Доменные (бизнес-) типы, которые инструменты возвращают наружу.
 * Это «причёсанное» представление поверх сырых сущностей OData —
 * именно его видит модель/пользователь, без GUID-шумихи и технических полей.
 *
 * Источник правды — Zod-схемы в src/schemas/output.ts (они же outputSchema
 * инструментов, SDK валидирует structuredContent против них в рантайме) —
 * типы выводятся из схем, а не дублируются вручную, чтобы не разъезжались.
 */

import type { z } from "zod";
import type { counterpartySchema, documentSummarySchema } from "../schemas/output.js";

export type Counterparty = z.infer<typeof counterpartySchema>;
export type DocumentSummary = z.infer<typeof documentSummarySchema>;
