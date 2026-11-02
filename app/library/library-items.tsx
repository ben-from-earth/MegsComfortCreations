'use client';

import LibraryItem from '@/library/library-item';
import LibraryPager from '@/library/library-pager';
import { useLibraryPageContext } from 'lib/context/library-page-context';
import { MEDIA_TYPE_DEFINITIONS } from 'lib/constants/media-types';
import { BOOK_SORT_SELECT_OPTIONS } from 'lib/constants/database-sort-options';

export default function LibraryItems() {
  const { libraryList, sortBy, ascDesc } = useLibraryPageContext();
  const { type, items, total, min, max } = libraryList;
  const plural = (
    MEDIA_TYPE_DEFINITIONS.find(({ mediaType }) => mediaType === type)
      ?.plural ?? 'items'
  ).toLowerCase();
  const sortLabel = (
    BOOK_SORT_SELECT_OPTIONS.find(({ value }) => value === sortBy)?.label ??
    'Title'
  ).toLowerCase();

  return (
    <section className="flex flex-col gap-4.5">
      <div className="flex flex-wrap items-center justify-between gap-4">
        <p className="m-0 font-semibold text-[#6b645d]">
          {total === 0
            ? `Showing 0 ${plural}`
            : `Showing ${min}–${max} of ${total} ${plural}`}{' '}
          · sorted by {sortLabel}, {ascDesc === 'asc' ? 'A–Z' : 'Z–A'}
        </p>
        <LibraryPager />
      </div>
      <div className="grid grid-cols-[repeat(auto-fill,minmax(min(380px,100%),1fr))] items-start gap-5">
        {items.map((item) => (
          <LibraryItem key={`${type}-${item.id}`} info={item} />
        ))}
      </div>
      {total > 0 ? (
        <div className="flex justify-end">
          <LibraryPager />
        </div>
      ) : null}
    </section>
  );
}
