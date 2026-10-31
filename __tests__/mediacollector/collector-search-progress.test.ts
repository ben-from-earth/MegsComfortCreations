import { collectorSearchProgressPercent } from '@/mediacollector/collector-search-progress';

describe('collectorSearchProgressPercent', () => {
  test('starts the collect phase at 0', () => {
    expect(
      collectorSearchProgressPercent({
        phase: 'collecting',
        searchedCount: 0,
        totalCount: 12,
      }),
    ).toBe(0);
  });

  test('uses 80 * (searched / total) during collection', () => {
    expect(
      collectorSearchProgressPercent({
        phase: 'collecting',
        searchedCount: 3,
        totalCount: 12,
      }),
    ).toBe(20);
  });

  test('reaches 80 when every title has been searched', () => {
    expect(
      collectorSearchProgressPercent({
        phase: 'collecting',
        searchedCount: 12,
        totalCount: 12,
      }),
    ).toBe(80);
  });

  test('holds at 80 while picking colors', () => {
    expect(
      collectorSearchProgressPercent({
        phase: 'pickingColors',
        isComplete: false,
      }),
    ).toBe(80);
  });

  test('finishes at 100 when color picking is done', () => {
    expect(
      collectorSearchProgressPercent({
        phase: 'pickingColors',
        isComplete: true,
      }),
    ).toBe(100);
  });

  test('returns 0 when there are no titles to search', () => {
    expect(
      collectorSearchProgressPercent({
        phase: 'collecting',
        searchedCount: 0,
        totalCount: 0,
      }),
    ).toBe(0);
  });
});
