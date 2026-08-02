type QueryMethodName =
  | 'select'
  | 'from'
  | 'where'
  | 'leftJoin'
  | 'innerJoin'
  | 'rightJoin'
  | 'orderBy'
  | 'groupBy'
  | 'limit'
  | 'offset'
  | 'returning'
  | 'values'
  | 'set'
  | 'onConflictDoNothing'
  | 'onConflictDoUpdate';

const QUERY_METHOD_NAMES = [
  'select',
  'from',
  'where',
  'leftJoin',
  'innerJoin',
  'rightJoin',
  'orderBy',
  'groupBy',
  'limit',
  'offset',
  'returning',
  'values',
  'set',
  'onConflictDoNothing',
  'onConflictDoUpdate',
] as const satisfies readonly QueryMethodName[];

export type ChainableQueryMock<T> = {
  [MethodName in QueryMethodName]: jest.Mock;
} & PromiseLike<T>;

/**
 * Fluent Drizzle-style chain that resolves when awaited at any depth.
 * Avoids brittle nested leftJoin/where/orderBy mock trees.
 */
export function createChainableQueryMock<T>(result: T): ChainableQueryMock<T> {
  const queryMock: Record<string, unknown> = {};

  for (const methodName of QUERY_METHOD_NAMES) {
    queryMock[methodName] = jest.fn(() => queryMock);
  }

  queryMock.then = (
    onFulfilled: (value: T) => unknown,
    onRejected?: (reason: unknown) => unknown,
  ) => Promise.resolve(result).then(onFulfilled, onRejected);

  return queryMock as unknown as ChainableQueryMock<T>;
}

export type MockDb = {
  select: jest.Mock;
  insert: jest.Mock;
  update: jest.Mock;
  delete: jest.Mock;
  transaction: jest.Mock;
};

export function createMockDb(options?: {
  selectResults?: unknown[];
  insertResult?: unknown;
  updateResult?: unknown;
  deleteResult?: unknown;
}): MockDb {
  const mockDb: MockDb = {
    select: jest.fn(() => createChainableQueryMock([])),
    insert: jest.fn(() =>
      createChainableQueryMock(
        options?.insertResult === undefined ? [] : options.insertResult,
      ),
    ),
    update: jest.fn(() =>
      createChainableQueryMock(
        options?.updateResult === undefined ? [] : options.updateResult,
      ),
    ),
    delete: jest.fn(() =>
      createChainableQueryMock(
        options?.deleteResult === undefined ? [] : options.deleteResult,
      ),
    ),
    transaction: jest.fn(),
  };

  mockDb.transaction.mockImplementation(
    async (callback: (transactionClient: MockDb) => unknown) =>
      callback(mockDb),
  );

  if (options?.selectResults) {
    for (const selectResult of options.selectResults) {
      mockDb.select.mockReturnValueOnce(createChainableQueryMock(selectResult));
    }
  }

  return mockDb;
}
