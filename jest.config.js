/** @type {import('jest').Config} */
module.exports = {
  preset: "ts-jest",
  testEnvironment: "node",
  roots: ["<rootDir>/tests"],
  testMatch: ["**/*.test.ts"],
  transform: { "^.+\\.tsx?$": ["ts-jest", { tsconfig: "tsconfig.json" }] },
  moduleNameMapper: {
    "^../src/(.*)$": "<rootDir>/src/$1",
  },
  collectCoverageFrom: ["src/**/*.ts", "!src/**/*.d.ts"],
  coverageThreshold: {
    global: { branches: 60, functions: 60, lines: 60, statements: 60 },
  },
};
