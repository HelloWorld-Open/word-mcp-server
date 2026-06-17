import { defineConfig } from "vitest/config"

export default defineConfig({
  test: {
    globals: true,
    environment: "node",
    include: ["tests/**/*.test.ts"],
    coverage: {
      provider: "v8",
      reporter: ["text", "lcov", "html"],
      include: ["src/**/*.ts"],
      exclude: [
        "src/word/com-proxy/com-proxy.mock.ts",
        "src/index.ts",
        "src/parent.ts",
        "src/child.ts",
      ],
    },
  },
})
