import { z } from "zod";

const ConfigSchema = z.object({
  PORT: z.coerce.number().int().positive().default(4000),
  WEB_ORIGIN: z.string().default("http://localhost:3000"),
  PLANNER_USE_DEMO_DATA: z.string().default("true").transform((value) => value === "true"),
  PLANNER_TIME_ZONE: z.string().default("America/New_York"),
  PLANNER_WORKDAY_START: z.string().regex(/^\d{2}:\d{2}$/).default("09:00"),
  PLANNER_WORKDAY_END: z.string().regex(/^\d{2}:\d{2}$/).default("17:00"),
  JIRA_BASE_URL: z.string().url().optional(),
  JIRA_EMAIL: z.string().email().optional(),
  JIRA_API_TOKEN: z.string().min(1).optional(),
  MICROSOFT_TENANT_ID: z.string().min(1).optional(),
  MICROSOFT_CLIENT_ID: z.string().min(1).optional(),
  MICROSOFT_CLIENT_SECRET: z.string().min(1).optional(),
  MICROSOFT_USER_ID: z.string().min(1).optional()
});

export type AppConfig = z.infer<typeof ConfigSchema>;

export function readConfig(environment: NodeJS.ProcessEnv = process.env): AppConfig {
  return ConfigSchema.parse(environment);
}
