import {
  codeBarColor,
  spotifyCodeUrl,
} from 'lib/trpc/routers/collect/actions/spotify-code';

const albumUri = 'spotify:album:rumours';

describe('codeBarColor', () => {
  test('uses black bars on a light color and white bars on a dark color', () => {
    expect(codeBarColor('#ffffff')).toBe('black');
    expect(codeBarColor('#000000')).toBe('white');
    expect(codeBarColor('#979797')).toBe('black');
    expect(codeBarColor('#969696')).toBe('white');
  });
});

describe('spotifyCodeUrl', () => {
  test('builds two png urls that share the album uri and differ in the background', () => {
    const light = spotifyCodeUrl(albumUri, '#ffffff');
    const dark = spotifyCodeUrl(albumUri, '#112233');

    expect(light).toBe(
      'https://scannables.scdn.co/uri/plain/png/ffffff/black/640/spotify:album:rumours',
    );
    expect(dark).toBe(
      'https://scannables.scdn.co/uri/plain/png/112233/white/640/spotify:album:rumours',
    );
    expect(light).toContain(albumUri);
    expect(dark).toContain(albumUri);
  });

  test('returns no url when the uri is empty or the color is not #rrggbb', () => {
    expect(spotifyCodeUrl('', '#ffffff')).toBeNull();
    expect(spotifyCodeUrl('   ', '#ffffff')).toBeNull();
    expect(spotifyCodeUrl(albumUri, '')).toBeNull();
    expect(spotifyCodeUrl(albumUri, '#fff')).toBeNull();
    expect(spotifyCodeUrl(albumUri, 'ffffff')).toBeNull();
  });
});
