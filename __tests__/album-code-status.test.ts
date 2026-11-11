import {
  albumCodeStatus,
  type AlbumCodeState,
} from 'lib/helpers/album-code-status';

const rumours: AlbumCodeState = {
  title: 'Rumours',
  artist: 'Fleetwood Mac',
  codeColor: '#e8dcc4',
  spotifyAlbumUri: null,
  codeStatus: undefined,
  codeGeneratedFor: null,
};

const generated: AlbumCodeState = {
  ...rumours,
  spotifyAlbumUri: 'spotify:album:rumours',
  codeStatus: 'generated',
  codeGeneratedFor: { title: 'Rumours', artist: 'Fleetwood Mac' },
};

describe('albumCodeStatus', () => {
  test('asks for a title and artist before a color', () => {
    expect(albumCodeStatus({ ...rumours, artist: '  ', codeColor: null })).toBe(
      'missingName',
    );
  });

  test('asks for a code color once the names are present', () => {
    expect(albumCodeStatus({ ...rumours, codeColor: null })).toBe(
      'missingColor',
    );
  });

  test('is ready with names and a color, including a collected uri', () => {
    expect(albumCodeStatus(rumours)).toBe('ready');
    expect(albumCodeStatus({ ...generated, codeStatus: undefined })).toBe(
      'ready',
    );
  });

  test('is generated for the stored title and artist', () => {
    expect(albumCodeStatus(generated)).toBe('generated');
    expect(albumCodeStatus({ ...generated, codeColor: null })).toBe(
      'generatedWithoutColor',
    );
  });

  test('is out of date after the title changes', () => {
    expect(albumCodeStatus({ ...generated, title: 'Tusk' })).toBe('outOfDate');
  });

  test('returns a no-match to ready when the title or artist is edited', () => {
    const noMatch: AlbumCodeState = {
      ...rumours,
      codeStatus: 'noMatch',
      codeGeneratedFor: { title: 'Rumours', artist: 'Fleetwood Mac' },
    };
    expect(albumCodeStatus(noMatch)).toBe('noMatch');
    expect(albumCodeStatus({ ...noMatch, artist: 'Fleetwood Mac ' })).toBe(
      'noMatch',
    );
    expect(albumCodeStatus({ ...noMatch, artist: 'Stevie Nicks' })).toBe(
      'ready',
    );
  });
});
