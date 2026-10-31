export const COLLECT_PHASE_MAX_PERCENT = 80;
export const COLOR_PHASE_PERCENT = 20;

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

export function collectorSearchProgressPercent(
  progress: CollectorSearchProgress,
): number {
  if (progress.phase === 'pickingColors') {
    return progress.isComplete
      ? COLLECT_PHASE_MAX_PERCENT + COLOR_PHASE_PERCENT
      : COLLECT_PHASE_MAX_PERCENT;
  }

  if (progress.totalCount <= 0) {
    return 0;
  }

  const searchedCount = Math.min(
    Math.max(progress.searchedCount, 0),
    progress.totalCount,
  );
  return COLLECT_PHASE_MAX_PERCENT * (searchedCount / progress.totalCount);
}
