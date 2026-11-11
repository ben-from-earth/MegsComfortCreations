export const ALBUM_CODE_RESULTS = ['generated', 'noMatch'] as const;
export type AlbumCodeResult = (typeof ALBUM_CODE_RESULTS)[number];

export type AlbumCodeStatus =
  | 'missingName'
  | 'missingColor'
  | 'ready'
  | 'generating'
  | 'generated'
  | 'generatedWithoutColor'
  | 'outOfDate'
  | 'noMatch';

export interface AlbumCodePair {
  title: string;
  artist: string;
}

export interface AlbumCodeState {
  title: string;
  artist: string | null | undefined;
  codeColor: string | null | undefined;
  spotifyAlbumUri: string | null | undefined;
  codeStatus: AlbumCodeResult | undefined;
  codeGeneratedFor: AlbumCodePair | null | undefined;
  isGenerating?: boolean;
}

export function albumCodePair(
  title: string,
  artist: string | null | undefined,
): AlbumCodePair {
  return { title: title.trim(), artist: (artist ?? '').trim() };
}

export function isSameAlbumCodePair(
  left: AlbumCodePair | null | undefined,
  right: AlbumCodePair,
) {
  return (
    left !== null &&
    left !== undefined &&
    left.title.trim() === right.title &&
    left.artist.trim() === right.artist
  );
}

export function albumCodeStatus(state: AlbumCodeState): AlbumCodeStatus {
  if (state.isGenerating) {
    return 'generating';
  }
  const pair = albumCodePair(state.title, state.artist);
  if (pair.title === '' || pair.artist === '') {
    return 'missingName';
  }
  const matchesLastLookup = isSameAlbumCodePair(state.codeGeneratedFor, pair);
  if (state.codeStatus === 'generated' && state.spotifyAlbumUri) {
    if (!matchesLastLookup) {
      return 'outOfDate';
    }
    return state.codeColor ? 'generated' : 'generatedWithoutColor';
  }
  if (state.codeStatus === 'noMatch' && matchesLastLookup) {
    return 'noMatch';
  }
  return state.codeColor ? 'ready' : 'missingColor';
}
