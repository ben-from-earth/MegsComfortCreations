import {
  defaultCoverRibbonVariant,
  defaultCoverTileLabel,
} from '@/components/shared/default-cover-ribbon';

describe('defaultCoverRibbonVariant', () => {
  it('shows no ribbon for a non-default cover', () => {
    expect(
      defaultCoverRibbonVariant({ url: 'a', isDefault: false }, 'a'),
    ).toBeNull();
  });

  it('marks the saved default and an unsaved new default', () => {
    expect(defaultCoverRibbonVariant({ url: 'a', isDefault: true }, 'a')).toBe(
      'saved',
    );
    expect(defaultCoverRibbonVariant({ url: 'b', isDefault: true }, 'a')).toBe(
      'new',
    );
    expect(
      defaultCoverRibbonVariant({ url: 'b', isDefault: true }, undefined),
    ).toBe('saved');
  });
});

describe('defaultCoverTileLabel', () => {
  it('describes the default state in the tile label', () => {
    expect(defaultCoverTileLabel(1, 'saved')).toBe('Cover option 1, default');
    expect(defaultCoverTileLabel(2, 'new')).toBe(
      'Cover option 2, new default, unsaved',
    );
    expect(defaultCoverTileLabel(3, null)).toBe('Cover option 3');
  });
});
