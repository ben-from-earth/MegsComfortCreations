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
      <div className='border-darkpink bg-lightpink relative flex min-w-lg flex-col items-center gap-2 self-center rounded-lg border-3 p-2 font-["Just_Another_Hand"] text-2xl tracking-wider shadow-[5px_5px_30px_rgba(0,0,0,0.3)]'>
        {items.map((item) => {
          return <LibraryItem key={`${type}-${item.id}`} info={item} />;
        })}
      </div>
    </section>
  );
}
