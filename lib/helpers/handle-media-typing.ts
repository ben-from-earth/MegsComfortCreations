import {
  type PostSavedMediaItem,
  type SavedBook,
} from '../interfaces/global-interfaces';

export function isBookRow(info: PostSavedMediaItem): info is SavedBook {
  return info.mediaType === 'book';
}
