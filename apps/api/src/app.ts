import cors from "@fastify/cors";
import Fastify from "fastify";
import {
  CalendarSyncRequestSchema,
  DailyRundownRequestSchema,
  DailyRundownSchema
} from "@workday/contracts";
import type { PlannerServices } from "./services.js";

export async function buildApp(services: PlannerServices, webOrigin = "http://localhost:3000") {
  const app = Fastify({ logger: false });
  await app.register(cors, { origin: webOrigin });

  app.get("/health", async () => ({ status: "ok" }));

  app.post("/api/rundown/daily", async (request, reply) => {
    const parsed = DailyRundownRequestSchema.safeParse(request.body ?? {});
    if (!parsed.success) return reply.status(400).send({ error: parsed.error.flatten() });
    const rundown = await services.createRundown(parsed.data.date);
    return DailyRundownSchema.parse(rundown);
  });

  app.post("/api/calendar/sync", async (request, reply) => {
    const parsed = CalendarSyncRequestSchema.safeParse(request.body);
    if (!parsed.success) return reply.status(400).send({ error: parsed.error.flatten() });
    return services.syncCalendar(parsed.data.schedule);
  });

  return app;
}
