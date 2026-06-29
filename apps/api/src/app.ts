import cors from "@fastify/cors";
import { timingSafeEqual } from "node:crypto";
import Fastify from "fastify";
import {
  CalendarSyncRequestSchema,
  DailyRundownRequestSchema,
  DailyRundownSchema,
  MicrosoftAuthStatusSchema,
  MicrosoftDeviceLoginSchema,
  MicrosoftDeviceLoginStatusSchema
} from "@workday/contracts";
import type { PlannerServices } from "./services.js";

export async function buildApp(
  services: PlannerServices,
  webOrigin = "http://localhost:3000",
  automationToken?: string
) {
  const app = Fastify({ logger: false });
  await app.register(cors, { origin: webOrigin });

  app.get("/health", async () => ({ status: "ok" }));

  app.get("/api/auth/microsoft/status", async () =>
    MicrosoftAuthStatusSchema.parse(await services.microsoftAuth.getStatus())
  );

  app.post("/api/auth/microsoft/device/start", async () =>
    MicrosoftDeviceLoginSchema.parse(await services.microsoftAuth.startDeviceLogin())
  );

  app.get<{ Params: { sessionId: string } }>(
    "/api/auth/microsoft/device/:sessionId",
    async (request) =>
      MicrosoftDeviceLoginStatusSchema.parse(
        services.microsoftAuth.getDeviceLoginStatus(request.params.sessionId)
      )
  );

  app.post("/api/auth/microsoft/disconnect", async () => {
    await services.microsoftAuth.disconnect();
    return { status: "disconnected" };
  });

  app.post("/api/rundown/daily", async (request, reply) => {
    const parsed = DailyRundownRequestSchema.safeParse(request.body ?? {});
    if (!parsed.success) return reply.status(400).send({ error: parsed.error.flatten() });
    const rundown = await services.createRundown(parsed.data);
    return DailyRundownSchema.parse(rundown);
  });

  app.post("/api/automation/daily", async (request, reply) => {
    if (!automationToken) {
      return reply.status(503).send({ error: "Automated triggers are not configured." });
    }
    const token = request.headers.authorization?.replace(/^Bearer\s+/i, "") ?? "";
    if (!tokensMatch(token, automationToken)) {
      return reply.status(401).send({ error: "Invalid automation token." });
    }
    const parsed = DailyRundownRequestSchema.safeParse(request.body ?? {});
    if (!parsed.success) return reply.status(400).send({ error: parsed.error.flatten() });
    return DailyRundownSchema.parse(await services.createRundown(parsed.data));
  });

  app.post("/api/calendar/sync", async (request, reply) => {
    const parsed = CalendarSyncRequestSchema.safeParse(request.body);
    if (!parsed.success) return reply.status(400).send({ error: parsed.error.flatten() });
    return services.syncCalendar(parsed.data.schedule);
  });

  return app;
}

function tokensMatch(candidate: string, expected: string): boolean {
  const candidateBuffer = Buffer.from(candidate);
  const expectedBuffer = Buffer.from(expected);
  return (
    candidateBuffer.length === expectedBuffer.length &&
    timingSafeEqual(candidateBuffer, expectedBuffer)
  );
}
