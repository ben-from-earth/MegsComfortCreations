import axios from 'axios';

export async function getOpenLibraryData(title: string, author?: string) {
  if (!author) {
    return { title, author, pubYear: null, pageCount: null };
  }

  const params = new URLSearchParams({
    title,
    author,
    limit: '1',
    fields: 'first_publish_year,number_of_pages_median',
  });
  const openLibraryRes = await axios.get(
    `https://openlibrary.org/search.json?${params.toString()}`,
  );
  const doc = openLibraryRes.data?.docs?.[0];
  if (!doc) {
    return { title, author, pubYear: null, pageCount: null };
  }
  const {
    first_publish_year: pubYear,
    number_of_pages_median: pageCount,
  }: { first_publish_year: number; number_of_pages_median: number } = doc;
  return { title, author, pubYear, pageCount };
}
