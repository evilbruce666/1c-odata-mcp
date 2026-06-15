/**
 * Типы протокола OData (1С реализует версию 3.0, JSON-формат)
 * и разобранной структуры $metadata (EDMX).
 */

/** Класс объекта 1С, определяемый по префиксу имени EntitySet. */
export type EntityClass =
  | "catalog" // Catalog_*          — справочники
  | "document" // Document_*         — документы
  | "accumulationRegister" // AccumulationRegister_*
  | "accountingRegister" // AccountingRegister_*
  | "informationRegister" // InformationRegister_*
  | "calculationRegister" // CalculationRegister_*
  | "enum" // Enum_*             — перечисления
  | "chartOfAccounts" // ChartOfAccounts_*
  | "chartOfCharacteristicTypes"
  | "constant" // Constant_*
  | "businessProcess"
  | "task"
  | "exchangePlan"
  | "documentJournal"
  | "other";

/** Ответ-коллекция OData в JSON (1С v3): значения в поле `value`. */
export interface ODataCollection<T> {
  "odata.metadata"?: string;
  "odata.count"?: string;
  value: T[];
}

/** Произвольная сущность OData: ключ Ref_Key + динамические поля. */
export interface ODataEntity {
  Ref_Key?: string;
  [field: string]: unknown;
}

/** Описание свойства (поля) сущности из $metadata. */
export interface MetaProperty {
  name: string;
  /** EDM-тип, напр. Edm.String, Edm.Decimal, Edm.DateTime, Edm.Guid, Edm.Boolean. */
  type: string;
  nullable: boolean;
}

/** Навигационное свойство (связь) из $metadata. */
export interface MetaNavigation {
  name: string;
  /** Целевой тип связи (имя EntityType). */
  toType: string;
  /** true, если связь «к многим». */
  collection: boolean;
}

/** Разобранное описание одной сущности. */
export interface EntityMeta {
  /** Техническое имя EntitySet, напр. "Catalog_Контрагенты". */
  entitySet: string;
  /** Имя EntityType. */
  entityType: string;
  /** Класс по префиксу. */
  class: EntityClass;
  /** «Человеческая» часть имени, напр. "Контрагенты". */
  shortName: string;
  keys: string[];
  properties: MetaProperty[];
  navigations: MetaNavigation[];
}

/** Полная карта сущностей базы, построенная из $metadata. */
export interface MetadataMap {
  /** Версия OData (напр. "3.0"). */
  odataVersion: string;
  /** Все сущности по имени EntitySet. */
  entities: Map<string, EntityMeta>;
}
