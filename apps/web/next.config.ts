import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  transpilePackages: ["@workday/contracts"]
};

export default nextConfig;
