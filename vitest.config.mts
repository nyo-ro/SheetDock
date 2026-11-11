import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // Keep ignored local backups and generated output outside the test suite.
    include: ["tests/**/*.test.{ts,tsx}"],
  },
});
