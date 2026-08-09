import type { MediaType } from 'lib/constants/mediaTypes';

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

export interface BlockInfo {
  title: string;
  author?: string | null;
  pubYear?: number | null;
  pageCount?: number | null;
  spineColor: string;
  genres?: string[];
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
