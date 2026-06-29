import { defineConfig } from "vitest/config";
import { fileURLToPath } from "node:url";

export default defineConfig({
  resolve: {
    alias: {
      "@workday/contracts": fileURLToPath(
        new URL("./packages/contracts/src/index.ts", import.meta.url)
      ),
      "@workday/planner": fileURLToPath(
        new URL("./packages/planner/src/index.ts", import.meta.url)
      ),
      "@workday/integrations": fileURLToPath(
        new URL("./packages/integrations/src/index.ts", import.meta.url)
      )
    }
  },
  test: {
    include: ["packages/**/*.test.ts", "apps/**/*.test.ts"],
    coverage: {
      reporter: ["text", "json", "html"]
    }
  }
});
