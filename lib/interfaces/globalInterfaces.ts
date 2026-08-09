import type { InferSelectModel, InferInsertModel } from 'drizzle-orm';
import { books, otherMedia } from '@/db/schema';
import type { MediaType, OtherMediaType } from 'lib/constants/mediaTypes';

export type MediaLabel = 'Book' | 'Movie' | 'Video Game' | 'Album';

export type BookRow = InferSelectModel<typeof books>;
export type OtherMediaRow = InferSelectModel<typeof otherMedia> & {
  mediaType: OtherMediaType;
};

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

export type PreSavedMediaItem = BookInsert | OtherMediaInsert;

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
  bookID: string;
  genres: string[];
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

export interface BlockInfo {
  title: string;
  author?: string | null;
  pubYear?: number | null;
  pageCount?: number | null;
  spineColor: string;
  genres?: string[];
  databaseGenres?: string[];
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
