import nextJest from 'next/jest.js';

const createJestConfig = nextJest({ dir: './' });

const config = {
  testEnvironment: 'node',
  setupFiles: ['<rootDir>/__tests__/setupEnv.ts'],
  testMatch: ['**/__tests__/**/*.test.ts', '**/__tests__/**/*.test.tsx'],
  moduleNameMapper: {
    '^@/(.*)$': '<rootDir>/app/$1',
    '^lib/(.*)$': '<rootDir>/lib/$1',
    '^superjson$': '<rootDir>/__tests__/mocks/superjson.ts',
  },
};

export default createJestConfig(config);
