export type CollectorSearchProgress =
  | {
      phase: 'collecting';
      searchedCount: number;
      totalCount: number;
    }
  | {
      phase: 'pickingColors';
      isComplete: boolean;
    };
