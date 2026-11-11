const CODE_COLOR = /^#([0-9a-fA-F]{6})$/;
const SCANNABLES_PNG_URL = 'https://scannables.scdn.co/uri/plain/png';

function codeColorChannels(hex: string) {
  const match = CODE_COLOR.exec(hex);
  if (!match) {
    return null;
  }
  const value = match[1];
  return {
    red: Number.parseInt(value.slice(0, 2), 16),
    green: Number.parseInt(value.slice(2, 4), 16),
    blue: Number.parseInt(value.slice(4, 6), 16),
    background: value.toLowerCase(),
  };
}

export function codeBarColor(hex: string): 'black' | 'white' {
  const channels = codeColorChannels(hex);
  if (!channels) {
    return 'white';
  }
  const luminance =
    (channels.red * 299 + channels.green * 587 + channels.blue * 114) / 1000;
  return luminance > 150 ? 'black' : 'white';
}

export function spotifyCodeUrl(
  spotifyAlbumUri: string,
  codeColor: string,
): string | null {
  const uri = spotifyAlbumUri.trim();
  const channels = codeColorChannels(codeColor);
  if (!uri || !channels) {
    return null;
  }
  return `${SCANNABLES_PNG_URL}/${channels.background}/${codeBarColor(codeColor)}/640/${uri}`;
}
