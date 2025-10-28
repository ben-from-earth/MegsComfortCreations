import { DatabaseSaveEditErrorResponse } from '@/app/api/api-Errors';

export type MediaType = 'book' | 'movie' | 'video_game' | 'album';

export type MediaLabel = 'Book' | 'Movie' | 'Video Game' | 'Album';

export interface presavedMediaItem {
  title: string;
  spine_color: string;
  blockID?: string;
  author?: string;
  pub_year?: number;
  page_count?: number;
  genres?: string[];
  image_urls: string[];
}

export interface postSavedMediaItem extends presavedMediaItem {
  id: string;
}

export interface SuccessfulMediaSearchResponse {
  message: string;
  foundMediaList: postSavedMediaItem[];
  total: number;
}

export interface SuccessfulMediaSaveEditResponse {
  message: string;
  actionAttemptItem: postSavedMediaItem;
  type: MediaType;
}

export interface SuccessfulPaginationResponse {
  message: string;
  paginatedList: postSavedMediaItem[];
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

export interface blockInfo {
  title: string;
  author?: string;
  pub_year?: number;
  page_count?: number;
  spine_color?: string;
  databaseGenres?: string[];
}

export interface Window {
  EyeDropper?: {
    new (): {
      open: () => Promise<{ sRGBHex: string }>;
    };
  };
}

export type databaseSaveServerResponse = (
  | DatabaseSaveEditErrorResponse
  | SuccessfulMediaSaveEditResponse
  | (SuccessfulMediaSaveEditResponse & {
      genreResponses: SuccessfulGenreLinkUnlinkResponse[];
    })
)[];
