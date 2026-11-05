import { MediaType } from 'lib/constants/media-types';
import axios from 'axios';

const API_KEY = process.env.GOOGLE_SEARCH_API_KEY;
const CX = process.env.GOOGLE_SEARCH_CX;

export function mediaCoverSearchQuery(
  title: string,
  author: string | undefined,
  type: MediaType,
) {
  if (type === 'movie') {
    return `${title} movie full dvd cover`;
  }
  return `${title}${author ? ` ${author}` : ''} ${type} Cover Image`;
}

export function mediaCoverResultCount(type: MediaType) {
  return type === 'movie' ? 5 : 3;
}

export async function getMediaCovers(
  title: string,
  author: string | undefined,
  type: MediaType,
) {
  const imgArr: string[] = [];
  if (!CX || !API_KEY) {
    throw new Error('Google Custom Search API key or CX not set');
  }
  const resultCount = mediaCoverResultCount(type);
  const params = new URLSearchParams({
    q: mediaCoverSearchQuery(title, author, type),
    cx: CX,
    key: API_KEY,
    searchType: 'image',
    num: String(resultCount),
  });
  const { data } = await axios.get<{ items?: { link: string }[] }>(
    `https://www.googleapis.com/customsearch/v1?${params.toString()}`,
  );
  (data.items ?? []).forEach((i) => imgArr.push(i.link));
  const successfulImages = imgArr.length;
  const missing = resultCount - successfulImages;
  for (let i = 0; i < missing; i++) {
    imgArr.push('/images/placeholder-image.png');
  }
  return imgArr;
}
