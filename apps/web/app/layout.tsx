import type { Metadata } from "next";
import "./styles.css";

export const metadata: Metadata = {
  title: "Workday Planner",
  description: "AI-assisted daily planning for Jira, Outlook, and Teams."
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
