import { z } from "zod";

const optionalString = (schema: z.ZodString = z.string()) =>
  z.preprocess((value) => (value === "" ? undefined : value), schema.optional());

const ConfigSchema = z.object({
  PORT: z.coerce.number().int().positive().default(4000),
  WEB_ORIGIN: z.string().default("http://localhost:3000"),
  PLANNER_USE_DEMO_DATA: z.string().default("true").transform((value) => value === "true"),
  PLANNER_TIME_ZONE: z.string().default("America/New_York"),
  PLANNER_WORKDAY_START: z.string().regex(/^\d{2}:\d{2}$/).default("09:00"),
  PLANNER_WORKDAY_END: z.string().regex(/^\d{2}:\d{2}$/).default("17:00"),
  JIRA_BASE_URL: optionalString(z.string().url()),
  JIRA_EMAIL: optionalString(z.string().email()),
  JIRA_API_TOKEN: optionalString(z.string().min(1)),
  MICROSOFT_AUTH_MODE: z
    .enum(["manual", "delegated", "client_credentials"])
    .default("manual"),
  MICROSOFT_TENANT_ID: z.string().min(1).default("organizations"),
  MICROSOFT_CLIENT_ID: optionalString(z.string().min(1)),
  MICROSOFT_CLIENT_SECRET: optionalString(z.string().min(1)),
  MICROSOFT_USER_ID: optionalString(z.string().min(1)),
  MICROSOFT_TOKEN_CACHE_PATH: optionalString(z.string().min(1)),
  MICROSOFT_ENABLE_TEAMS_CHANNELS: z
    .string()
    .default("false")
    .transform((value) => value === "true"),
  PLANNER_AUTOMATION_TOKEN: optionalString(z.string().min(16))
});

export type AppConfig = z.infer<typeof ConfigSchema>;

export function readConfig(environment: NodeJS.ProcessEnv = process.env): AppConfig {
  return ConfigSchema.parse(environment);
}
