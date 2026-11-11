import {
  pickSpotifyAlbum,
  spotifyAlbumSearchQuery,
  toAlbumCover,
  toReleaseDate,
} from 'lib/trpc/routers/collect/actions/get-spotify-album';

const rumours = {
  name: 'Rumours',
  release_date: '1977-02-04',
};
const rumoursLive = {
  name: 'Rumours Live',
  release_date: '2023-09-08',
};

describe('spotifyAlbumSearchQuery', () => {
  test('quotes the title and artist', () => {
    expect(spotifyAlbumSearchQuery('Rumours', 'Fleetwood Mac')).toBe(
      'album:"Rumours" artist:"Fleetwood Mac"',
    );
  });
});

describe('pickSpotifyAlbum', () => {
  test('takes the first album when the card has no full release date', () => {
    expect(pickSpotifyAlbum([rumoursLive, rumours])).toBe(rumoursLive);
    expect(pickSpotifyAlbum([rumoursLive, rumours], '1977')).toBe(rumoursLive);
    expect(pickSpotifyAlbum([])).toBeUndefined();
  });

  test('prefers the album whose Spotify date starts with the release year', () => {
    expect(pickSpotifyAlbum([rumoursLive, rumours], '1977-02-04')).toBe(rumours);
  });

  test('keeps the first album when no release year matches', () => {
    expect(pickSpotifyAlbum([rumoursLive, rumours], '1969-01-01')).toBe(
      rumoursLive,
    );
  });
});

describe('toAlbumCover', () => {
  test('uses the widest image Spotify listed first', () => {
    expect(
      toAlbumCover({
        images: [{ url: 'https://covers/wide.jpg' }, { url: 'https://covers/small.jpg' }],
      }),
    ).toBe('https://covers/wide.jpg');
  });

  test('returns null when the album has no artwork', () => {
    expect(toAlbumCover({ images: [] })).toBeNull();
    expect(toAlbumCover({ images: null })).toBeNull();
    expect(toAlbumCover({})).toBeNull();
  });
});

describe('toReleaseDate', () => {
  test('keeps a day-precision date and drops a year or a month', () => {
    expect(
      toReleaseDate({
        release_date: '1977-02-04',
        release_date_precision: 'day',
      }),
    ).toBe('1977-02-04');
    expect(
      toReleaseDate({
        release_date: '1977',
        release_date_precision: 'year',
      }),
    ).toBeNull();
    expect(
      toReleaseDate({
        release_date: '1977-02',
        release_date_precision: 'month',
      }),
    ).toBeNull();
  });
});
