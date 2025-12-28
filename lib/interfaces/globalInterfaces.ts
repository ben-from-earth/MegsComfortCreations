// drizzle types
import type { InferSelectModel, InferInsertModel } from 'drizzle-orm';
import { books, movies, videoGames, albums } from '@//db/schema';

import { DatabaseSaveEditErrorResponse } from '@//api/api-Errors';

export type MediaType = 'book' | 'movie' | 'videoGame' | 'album';
export type MediaLabel = 'Book' | 'Movie' | 'Video Game' | 'Album';

// 1. Map Drizzle row types
export type BookRow = InferSelectModel<typeof books>;
export type MovieRow = InferSelectModel<typeof movies>;
export type VideoGameRow = InferSelectModel<typeof videoGames>;
export type AlbumRow = InferSelectModel<typeof albums>;

// 2. Extras that don’t live in the DB but you still want on responses
interface MediaExtras {
  blockID?: string;
  genres?: string[];
}

export type BookInsert = Omit<InferInsertModel<typeof books>, 'id'> &
  MediaExtras;
export type MovieInsert = Omit<InferInsertModel<typeof movies>, 'id'> &
  MediaExtras;
export type VideoGameInsert = Omit<InferInsertModel<typeof videoGames>, 'id'> &
  MediaExtras;
export type AlbumInsert = Omit<InferInsertModel<typeof albums>, 'id'> &
  MediaExtras;

// 3. “Pre-saved” (create/edit body) – based on INSERT types, no id
//    You *can* keep title/spineColor etc. here if your JSON schema
//    is slightly different from the DB, but this keeps you close to Drizzle.
export type PreSavedMediaItem =
  | BookInsert
  | MovieInsert
  | VideoGameInsert
  | AlbumInsert;

// 4. “Post-saved” (what comes back from DB) – based on SELECT types
export type PostSavedMediaItem = BookRow | MovieRow | VideoGameRow | AlbumRow;

// 5. Response shapes now use these types

export interface SuccessfulMediaSearchResponse {
  message: string;
  foundMediaList: PostSavedMediaItem[];
  total: number;
}

export interface SuccessfulMediaSaveEditResponse {
  message: string;
  actionAttemptItem: PostSavedMediaItem & MediaExtras;
  type: MediaType;
}

export interface SuccessfulPaginationResponse {
  message: string;
  paginatedList: PostSavedMediaItem[];
  total: number;
}

export interface GenreLinkUnlinkRequest {
  bookID: string; // uuid (string in TS) from BookRow['id']
  genres: string[];
}

export interface SuccessfulGenreLinkUnlinkResponse {
  message: string;
  genre: string;
  bookID: string;
}

// 6. Make this match your camelCase DB schema (or keep snake_case if your rows do)
export interface BlockInfo {
  title: string;
  author?: string;
  pubYear?: number | null;
  pageCount?: number | null;
  spineColor?: string;
  databaseGenres?: string[];
}

// 7. Server response union
export type DatabaseSaveServerResponse = (
  | DatabaseSaveEditErrorResponse
  | SuccessfulMediaSaveEditResponse
  | (SuccessfulMediaSaveEditResponse & {
      genreResponses: SuccessfulGenreLinkUnlinkResponse[];
    })
)[];
