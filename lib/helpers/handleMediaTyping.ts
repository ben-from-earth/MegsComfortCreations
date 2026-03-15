import {
  PostSavedMediaItem,
} from '../interfaces/globalInterfaces';

export function isBookRow(
  info: PostSavedMediaItem,
): info is PostSavedMediaItem & {
  author: string;
  pageCount: number | null;
  pubYear: number | null;
} {
  return typeof info.author === 'string';
}
