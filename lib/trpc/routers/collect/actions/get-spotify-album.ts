const SPOTIFY_TOKEN_URL = 'https://accounts.spotify.com/api/token';
const SPOTIFY_SEARCH_URL = 'https://api.spotify.com/v1/search';
const TOKEN_REFRESH_BUFFER_MS = 60_000;
const FULL_RELEASE_DATE = /^\d{4}-\d{2}-\d{2}$/;

export type SpotifyCatalogAlbum = {
  uri: string;
  name: string;
  artists: { name: string }[];
  release_date: string;
  release_date_precision: string;
  images: { url: string }[] | null;
};

export type SpotifyAlbumResult =
  | {
      found: true;
      uri: string;
      name: string;
      artist: string;
      releaseDate: string | null;
      coverUrl: string | null;
    }
  | { found: false };

type CachedToken = {
  accessToken: string;
  expiresAt: number;
};

let cachedToken: CachedToken | null = null;

export function spotifyAlbumSearchQuery(title: string, artist: string) {
  return `album:"${title}" artist:"${artist}"`;
}

export function pickSpotifyAlbum<T extends { release_date: string }>(
  albums: T[],
  releaseDate?: string | null,
): T | undefined {
  const firstAlbum = albums[0];
  if (!firstAlbum) {
    return undefined;
  }

  const year =
    releaseDate && FULL_RELEASE_DATE.test(releaseDate)
      ? releaseDate.slice(0, 4)
      : null;
  if (!year) {
    return firstAlbum;
  }

  return (
    albums.find((album) => album.release_date.startsWith(year)) ?? firstAlbum
  );
}

export function toAlbumCover(album: {
  images?: { url: string }[] | null;
}): string | null {
  const coverUrl = album.images?.[0]?.url;
  return coverUrl ? coverUrl : null;
}

export function toReleaseDate(album: {
  release_date?: string | null;
  release_date_precision?: string | null;
}): string | null {
  if (
    album.release_date_precision === 'day' &&
    album.release_date &&
    FULL_RELEASE_DATE.test(album.release_date)
  ) {
    return album.release_date;
  }
  return null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function readSpotifyAlbum(value: unknown): SpotifyCatalogAlbum | null {
  if (!isRecord(value) || typeof value.name !== 'string') {
    return null;
  }

  const uri =
    typeof value.uri === 'string' && value.uri.startsWith('spotify:album:')
      ? value.uri
      : typeof value.id === 'string' && value.id.length > 0
        ? `spotify:album:${value.id}`
        : null;
  if (!uri) {
    return null;
  }

  const artists = Array.isArray(value.artists)
    ? value.artists.flatMap((artist) => {
        if (!isRecord(artist) || typeof artist.name !== 'string') {
          return [];
        }
        return [{ name: artist.name }];
      })
    : [];
  const images = Array.isArray(value.images)
    ? value.images.flatMap((image) => {
        if (!isRecord(image) || typeof image.url !== 'string' || !image.url) {
          return [];
        }
        return [{ url: image.url }];
      })
    : null;

  return {
    uri,
    name: value.name,
    artists,
    release_date: typeof value.release_date === 'string' ? value.release_date : '',
    release_date_precision:
      typeof value.release_date_precision === 'string'
        ? value.release_date_precision
        : '',
    images,
  };
}

function readAlbumPage(value: unknown): SpotifyCatalogAlbum[] {
  if (!isRecord(value) || !isRecord(value.albums) || !Array.isArray(value.albums.items)) {
    return [];
  }
  return value.albums.items.flatMap((item) => {
    const album = readSpotifyAlbum(item);
    return album ? [album] : [];
  });
}

function readToken(value: unknown): { accessToken: string; expiresIn: number } | null {
  if (!isRecord(value) || typeof value.access_token !== 'string' || !value.access_token) {
    return null;
  }
  return {
    accessToken: value.access_token,
    expiresIn: typeof value.expires_in === 'number' ? value.expires_in : 3600,
  };
}

async function readJson(response: Response): Promise<unknown> {
  try {
    return await response.json();
  } catch {
    return null;
  }
}

async function fetchAccessToken(): Promise<string | null> {
  if (cachedToken && cachedToken.expiresAt - Date.now() > TOKEN_REFRESH_BUFFER_MS) {
    return cachedToken.accessToken;
  }

  const clientId = process.env.SPOTIFY_CLIENT_ID?.trim();
  const clientSecret = process.env.SPOTIFY_CLIENT_SECRET?.trim();
  if (!clientId || !clientSecret) {
    return null;
  }

  const response = await fetch(SPOTIFY_TOKEN_URL, {
    method: 'POST',
    headers: {
      Authorization: `Basic ${Buffer.from(`${clientId}:${clientSecret}`).toString('base64')}`,
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: new URLSearchParams({ grant_type: 'client_credentials' }),
  });
  if (!response.ok) {
    return null;
  }

  const token = readToken(await readJson(response));
  if (!token) {
    return null;
  }

  cachedToken = {
    accessToken: token.accessToken,
    expiresAt: Date.now() + token.expiresIn * 1000,
  };
  return token.accessToken;
}

function toFoundAlbum(album: SpotifyCatalogAlbum): SpotifyAlbumResult {
  return {
    found: true,
    uri: album.uri,
    name: album.name,
    artist: album.artists.map((artist) => artist.name).join(', '),
    releaseDate: toReleaseDate(album),
    coverUrl: toAlbumCover(album),
  };
}

async function searchSpotifyAlbums(
  accessToken: string,
  title: string,
  artist: string,
) {
  const params = new URLSearchParams({
    q: spotifyAlbumSearchQuery(title, artist),
    type: 'album',
    limit: '10',
  });
  return fetch(`${SPOTIFY_SEARCH_URL}?${params.toString()}`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
}

export async function getSpotifyAlbum(
  title: string,
  artist: string,
  releaseDate?: string | null,
): Promise<SpotifyAlbumResult> {
  const albumTitle = title.trim();
  const albumArtist = artist.trim();
  if (!albumTitle || !albumArtist) {
    return { found: false };
  }

  try {
    const accessToken = await fetchAccessToken();
    if (!accessToken) {
      return { found: false };
    }

    let response = await searchSpotifyAlbums(accessToken, albumTitle, albumArtist);
    if (response.status === 401) {
      cachedToken = null;
      const refreshedToken = await fetchAccessToken();
      if (!refreshedToken) {
        return { found: false };
      }
      response = await searchSpotifyAlbums(refreshedToken, albumTitle, albumArtist);
    }
    if (!response.ok) {
      return { found: false };
    }

    const album = pickSpotifyAlbum(
      readAlbumPage(await readJson(response)),
      releaseDate,
    );
    return album ? toFoundAlbum(album) : { found: false };
  } catch {
    return { found: false };
  }
}
