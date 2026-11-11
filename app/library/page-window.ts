export type PageWindowItem = number | 'ellipsis';

export function getPageWindow(
  page: number,
  pageCount: number,
): PageWindowItem[] {
  if (pageCount < 1) {
    return [];
  }

  const includedPages = new Set<number>();
  for (const pageNumber of [1, pageCount, page - 1, page, page + 1]) {
    if (pageNumber >= 1 && pageNumber <= pageCount) {
      includedPages.add(pageNumber);
    }
  }

  const sortedPages = [...includedPages].sort((left, right) => left - right);
  const window: PageWindowItem[] = [];
  for (const [index, pageNumber] of sortedPages.entries()) {
    const previousPage = sortedPages[index - 1];
    if (index > 0 && pageNumber - previousPage > 1) {
      window.push('ellipsis');
    }
    window.push(pageNumber);
  }
  return window;
}
