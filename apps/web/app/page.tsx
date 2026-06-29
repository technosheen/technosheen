import { PlannerDashboard } from "../components/planner-dashboard";

export default function Home() {
  const timeZone = process.env.NEXT_PUBLIC_PLANNER_TIME_ZONE ?? "America/New_York";
  const dateLabel = new Intl.DateTimeFormat("en-US", {
    weekday: "long",
    month: "long",
    day: "numeric",
    timeZone
  }).format(new Date());

  return <PlannerDashboard dateLabel={dateLabel} />;
}
