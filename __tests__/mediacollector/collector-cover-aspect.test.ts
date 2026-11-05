import {
  collectorCoverAspectClassName,
  coverAspectClassName,
} from '@/mediacollector/cover-tile';

describe('collectorCoverAspectClassName', () => {
  test('uses a 3:2 frame for movie covers', () => {
    expect(collectorCoverAspectClassName('movie')).toBe('aspect-[3/2]');
  });

  test('leaves the other collector frames unchanged', () => {
    expect(collectorCoverAspectClassName('book')).toBe('aspect-[2/3]');
    expect(collectorCoverAspectClassName('album')).toBe('aspect-square');
    expect(collectorCoverAspectClassName('videoGame')).toBe(
      coverAspectClassName('videoGame'),
    );
  });
});
