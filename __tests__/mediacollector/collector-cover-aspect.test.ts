import { collectorCoverAspectClassName } from '@/mediacollector/cover-tile';

describe('collectorCoverAspectClassName', () => {
  test('uses a 3:2 frame for movie and video game covers', () => {
    expect(collectorCoverAspectClassName('movie')).toBe('aspect-[3/2]');
    expect(collectorCoverAspectClassName('videoGame')).toBe('aspect-[3/2]');
  });

  test('leaves the book and album frames unchanged', () => {
    expect(collectorCoverAspectClassName('book')).toBe('aspect-[2/3]');
    expect(collectorCoverAspectClassName('album')).toBe('aspect-square');
  });
});
