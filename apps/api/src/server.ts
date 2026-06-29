import { loadEnvFile } from "node:process";
import { fileURLToPath } from "node:url";
import { buildApp } from "./app.js";
import { readConfig } from "./config.js";
import { buildPlannerServices } from "./services.js";

try {
  loadEnvFile(fileURLToPath(new URL("../../../.env", import.meta.url)));
} catch (error) {
  if (!isMissingEnvFile(error)) throw error;
}

const config = readConfig();
const app = await buildApp(
  buildPlannerServices(config),
  config.WEB_ORIGIN,
  config.PLANNER_AUTOMATION_TOKEN
);

try {
  await app.listen({ port: config.PORT, host: "0.0.0.0" });
} catch (error) {
  app.log.error(error);
  process.exit(1);
}

function isMissingEnvFile(error: unknown): boolean {
  return (
    typeof error === "object" &&
    error !== null &&
    "code" in error &&
    error.code === "ENOENT"
  );
}
