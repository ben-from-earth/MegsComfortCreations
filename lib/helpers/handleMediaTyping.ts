import {
  BookRow,
  MediaType,
  PostSavedMediaItem,
} from '../interfaces/globalInterfaces';

export function isBookRow(
  type: MediaType,
  info: PostSavedMediaItem,
): info is BookRow {
  return type === 'book';
}
