import type { MediaType } from 'lib/constants/media-types';
import type { MovieFormat } from 'lib/constants/movie-formats';

interface MediaExtras {
  blockID?: string;
  genres?: string[];
}

export interface MediaImageItem {
  url: string;
  selected?: boolean;
  isDefault: boolean;
  spineColor: string;
  codeColor?: string | null;
}

interface SavedMediaBase {
  id: string;
  title: string;
  spineColor: string;
  images: MediaImageItem[];
  genres?: string[];
}

export interface SavedBook extends SavedMediaBase {
  mediaType: 'book';
  author: string;
  pageCount: number | null;
  pubYear: number | null;
}

export interface SavedMovie extends SavedMediaBase {
  mediaType: 'movie';
  movieFormat: MovieFormat;
}

export interface SavedVideoGame extends SavedMediaBase {
  mediaType: 'videoGame';
  console: string;
}

export interface SavedAlbum extends SavedMediaBase {
  mediaType: 'album';
  artist: string;
  releaseDate: string | null;
  spotifyAlbumUri: string | null;
}

export type PostSavedMediaItem =
  | SavedBook
  | SavedMovie
  | SavedVideoGame
  | SavedAlbum;

export interface SuccessfulPaginationResponse {
  message: string;
  paginatedList: PostSavedMediaItem[];
  total: number;
}

export interface SuccessfulGenreLinkUnlinkResponse {
  message: string;
  genre: string;
  bookID: string;
}

export interface DatabaseSaveEditErrorResponse {
  error: string;
  message: string;
  errors: string[];
  title: string;
}

export type DatabaseSaveSuccessResult = {
  success: true;
  blockID: string;
  title: string;
  message: string;
  type: MediaType;
  actionAttemptItem: PostSavedMediaItem & MediaExtras;
  genreResponses?: SuccessfulGenreLinkUnlinkResponse[];
};

export type DatabaseSaveFailureResult = {
  success: false;
  blockID: string;
  title: string;
  error: string;
  message: string;
  errors: string[];
};

export type DatabaseSaveResultItem =
  | DatabaseSaveSuccessResult
  | DatabaseSaveFailureResult;

export type DatabaseSaveServerResponse = DatabaseSaveResultItem[];
