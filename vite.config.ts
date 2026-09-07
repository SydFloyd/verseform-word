import { defineConfig } from "vitest/config";
import devCerts from "office-addin-dev-certs";

export default defineConfig(async ({ command }) => {
  const isTestRun = process.env.VITEST === "true";
  const server = command === "serve" && !isTestRun
    ? {
        host: "localhost" as const,
        port: 3000,
        strictPort: true,
        https: await devCerts.getHttpsServerOptions(),
      }
    : undefined;

  return {
    build: {
      target: "es2022",
      sourcemap: true,
    },
    server,
    // Native host proofs run the built assets. Vite's development client injects
    // inline CSS and opens a WebSocket, both intentionally denied by our CSP.
    preview: server,
    test: {
      environment: "node",
    },
  };
});
