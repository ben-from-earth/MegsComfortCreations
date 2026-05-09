// drizzle types
import type { InferSelectModel, InferInsertModel } from 'drizzle-orm';
import { books, otherMedia } from '@/db/schema';
import type { MediaType, OtherMediaType } from 'lib/constants/mediaTypes';

import { DatabaseSaveEditErrorResponse } from '@/api/api-Errors';

export type MediaLabel = 'Book' | 'Movie' | 'Video Game' | 'Album';

// 1. Map Drizzle row types
export type BookRow = InferSelectModel<typeof books>;
export type OtherMediaRow = InferSelectModel<typeof otherMedia> & {
  mediaType: OtherMediaType;
};

// 2. Extras that don’t live in the DB but you still want on responses
interface MediaExtras {
  blockID?: string;
  genres?: string[];
}

export interface MediaImageItem {
  url: string;
  selected?: boolean;
  isDefault: boolean;
  spineColor: string;
}

export type BookInsert = Omit<InferInsertModel<typeof books>, 'id'> &
  MediaExtras;
export type OtherMediaInsert = Omit<InferInsertModel<typeof otherMedia>, 'id'> &
  MediaExtras;

// 3. “Pre-saved” (create/edit body) – based on INSERT types, no id
//    You *can* keep title/spineColor etc. here if your JSON schema
//    is slightly different from the DB, but this keeps you close to Drizzle.
export type PreSavedMediaItem = BookInsert | OtherMediaInsert;

// 4. “Post-saved” (what comes back from DB)
export interface PostSavedMediaItem {
  id: string;
  title: string;
  spineColor: string;
  images: MediaImageItem[];
  mediaType?: string;
  author?: string;
  pageCount?: number | null;
  pubYear?: number | null;
}

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
  author?: string | null;
  pubYear?: number | null;
  pageCount?: number | null;
  spineColor: string;
  genres?: string[];
  databaseGenres?: string[];
}

// 7. Server response union
type DatabaseSaveServerResponseItem =
  | DatabaseSaveEditErrorResponse
  | SuccessfulMediaSaveEditResponse
  | (SuccessfulMediaSaveEditResponse & {
      genreResponses: SuccessfulGenreLinkUnlinkResponse[];
    });

export type DatabaseSaveServerResponse = DatabaseSaveServerResponseItem[];
