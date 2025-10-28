// jest.config.ts
import nextJest from "next/jest.js";

const createJestConfig = nextJest({ dir: "./" });

const config: import("jest").Config = {
  testEnvironment: "node",

  testMatch: ["**/__tests__/**/*.test.ts"],
  moduleNameMapper: {
    "^@/(.*)$": "<rootDir>/$1",
  },
};
export default createJestConfig(config);
