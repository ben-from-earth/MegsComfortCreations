import { MediaType } from 'lib/constants/mediaTypes';
import axios from 'axios';

const API_KEY = process.env.GOOGLE_SEARCH_API_KEY;
const CX = process.env.GOOGLE_SEARCH_CX;

export async function getMediaCovers(
  title: string,
  author: string | undefined,
  type: MediaType,
) {
  const imgArr: string[] = [];
  if (!CX || !API_KEY) {
    throw new Error('Google Custom Search API key or CX not set');
  }
  const params = new URLSearchParams({
    q: `${title}${author ? ` ${author}` : ''} ${type} Cover Image`,
    cx: CX,
    key: API_KEY,
    searchType: 'image',
    num: '3',
  });
  const { data } = await axios.get<{ items?: { link: string }[] }>(
    `https://www.googleapis.com/customsearch/v1?${params.toString()}`,
  );
  (data.items ?? []).forEach((i) => imgArr.push(i.link));
  const successfulImages = imgArr.length;
  const missing = 3 - successfulImages;
  for (let i = 0; i < missing; i++) {
    imgArr.push('/images/placeholder-image.png');
  }
  return imgArr;
}
