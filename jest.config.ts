// jest.config.ts
import nextJest from "next/jest.js";

const createJestConfig = nextJest({ dir: "./" });

const config: import("jest").Config = {
  testEnvironment: "node",
  setupFiles: ["<rootDir>/__tests__/setupEnv.ts"],

  testMatch: ["**/__tests__/**/*.test.ts"],
  moduleNameMapper: {
    "^@/(.*)$": "<rootDir>/app/$1",
    "^lib/(.*)$": "<rootDir>/lib/$1",
    "^superjson$": "<rootDir>/__tests__/mocks/superjson.ts",
  },
};
export default createJestConfig(config);
