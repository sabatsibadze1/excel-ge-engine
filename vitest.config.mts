import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "node",
    include: ["src/**/*.test.ts", "scripts/**/*.test.ts"],
    // HyperFormula instances take a few milliseconds each; the CLI and worker-build tests spawn Node.
    testTimeout: 60_000,
  },
});
