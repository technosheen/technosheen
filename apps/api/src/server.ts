import { buildApp } from "./app.js";
import { readConfig } from "./config.js";
import { buildPlannerServices } from "./services.js";

const config = readConfig();
const app = await buildApp(buildPlannerServices(config), config.WEB_ORIGIN);

try {
  await app.listen({ port: config.PORT, host: "0.0.0.0" });
} catch (error) {
  app.log.error(error);
  process.exit(1);
}
