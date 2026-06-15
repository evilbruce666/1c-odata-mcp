import type { AppConfig } from "./config/env.js";
import { ODataClient } from "./odata/client.js";
import { loadMetadata } from "./odata/metadata.js";
import type { MetadataMap } from "./types/odata.js";

/**
 * Общий контекст сервера: конфигурация, OData-клиент и лениво-кешируемая
 * карта метаданных. Инструменты получают его и не знают деталей транспорта.
 */
export class ServerContext {
  readonly cfg: AppConfig;
  readonly client: ODataClient;
  private metaPromise: Promise<MetadataMap> | undefined;

  constructor(cfg: AppConfig) {
    this.cfg = cfg;
    this.client = new ODataClient(cfg);
  }

  /** Карта сущностей из $metadata (загружается один раз, затем из кеша). */
  getMetadata(): Promise<MetadataMap> {
    if (!this.metaPromise) {
      this.metaPromise = loadMetadata(this.client).catch((e) => {
        // Сброс кеша при ошибке, чтобы следующий вызов попробовал снова.
        this.metaPromise = undefined;
        throw e;
      });
    }
    return this.metaPromise;
  }

  /** Множество доступных имён EntitySet — для resolveEntity(). */
  async available(): Promise<ReadonlySet<string>> {
    const meta = await this.getMetadata();
    return new Set(meta.entities.keys());
  }
}
