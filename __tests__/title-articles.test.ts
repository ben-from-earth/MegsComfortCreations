import {
  toDisplayTitle,
  toStoredTitle,
  toTitleSearchTerm,
} from 'lib/helpers/title-articles';

describe('toStoredTitle', () => {
  test.each([
    ['The Hobbit', 'Hobbit, The'],
    ['A Wrinkle in Time', 'Wrinkle in Time, A'],
    ['An Echo in the Bone', 'Echo in the Bone, An'],
    ['the hobbit', 'hobbit, the'],
  ])('moves the leading article of %p to the end', (title, stored) => {
    expect(toStoredTitle(title)).toBe(stored);
  });

  test('leaves an already-stored title alone', () => {
    expect(toStoredTitle('Hobbit, The')).toBe('Hobbit, The');
  });

  test.each(['Dune', 'Theodore Boone', 'Anne of Green Gables', 'Atonement'])(
    'leaves %p alone',
    (title) => {
      expect(toStoredTitle(title)).toBe(title);
    },
  );
});

describe('toDisplayTitle', () => {
  test.each([
    ['Hobbit, The', 'The Hobbit'],
    ['Wrinkle in Time, A', 'A Wrinkle in Time'],
    ['Echo in the Bone, An', 'An Echo in the Bone'],
  ])('moves the trailing article of %p to the front', (stored, title) => {
    expect(toDisplayTitle(stored)).toBe(title);
  });

  test('leaves an already-displayed title alone', () => {
    expect(toDisplayTitle('The Hobbit')).toBe('The Hobbit');
  });

  test('round-trips through the stored form', () => {
    expect(toDisplayTitle(toStoredTitle('The Name of the Wind'))).toBe(
      'The Name of the Wind',
    );
  });
});

describe('toTitleSearchTerm', () => {
  test.each([
    ['The Hob', 'Hob'],
    ['hobbit, the', 'hobbit'],
    ['Dune', 'Dune'],
  ])('searches %p as %p', (title, searchTerm) => {
    expect(toTitleSearchTerm(title)).toBe(searchTerm);
  });
});
