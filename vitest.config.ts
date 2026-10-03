import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // Тесты негативных сценариев пишут ошибки через логгер — не в рабочий лог пользователя.
    env: { LOG_LEVEL: "silent" },
  },
});
