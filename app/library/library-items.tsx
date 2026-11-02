'use client';

import { useCallback, useState, useTransition } from 'react';
import LibraryItem from '@/library/library-item';
import LibraryPager from '@/library/library-pager';
import { useLibraryPageContext } from 'lib/context/library-page-context';
import LibraryPendingSavesContext, {
  type RegisterPendingSave,
} from 'lib/context/library-pending-saves-context';
import { MEDIA_TYPE_DEFINITIONS } from 'lib/constants/media-types';
import { BOOK_SORT_SELECT_OPTIONS } from 'lib/constants/database-sort-options';

export default function LibraryItems() {
  const { libraryList, sortBy, ascDesc } = useLibraryPageContext();
  const { type, items, total, min, max } = libraryList;
  const [pendingSaves, setPendingSaves] = useState<
    ReadonlyMap<string, () => Promise<void>>
  >(() => new Map());
  const [isSavingAll, startSavingAll] = useTransition();
  const registerPendingSave = useCallback<RegisterPendingSave>(
    (blockID, save) => {
      setPendingSaves((current) => new Map(current).set(blockID, save));
      return () =>
        setPendingSaves((current) => {
          const next = new Map(current);
          next.delete(blockID);
          return next;
        });
    },
    [],
  );
  const plural = (
    MEDIA_TYPE_DEFINITIONS.find(({ mediaType }) => mediaType === type)
      ?.plural ?? 'items'
  ).toLowerCase();
  const sortLabel = (
    BOOK_SORT_SELECT_OPTIONS.find(({ value }) => value === sortBy)?.label ??
    'Title'
  ).toLowerCase();

  return (
    <LibraryPendingSavesContext value={registerPendingSave}>
      <section className="flex flex-col gap-4.5">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <p className="m-0 font-semibold text-[#6b645d]">
            {total === 0
              ? `Showing 0 ${plural}`
              : `Showing ${min}–${max} of ${total} ${plural}`}{' '}
            · sorted by {sortLabel}, {ascDesc === 'asc' ? 'A–Z' : 'Z–A'}
          </p>
          <div className="flex flex-wrap items-center gap-3">
            {pendingSaves.size > 1 ? (
              <button
                type="button"
                disabled={isSavingAll}
                onClick={() =>
                  startSavingAll(async () => {
                    await Promise.all(
                      [...pendingSaves.values()].map((save) => save()),
                    );
                  })
                }
                className="flex h-11 cursor-pointer items-center justify-center gap-2 rounded-xl border border-transparent bg-[#2a4637] px-5 text-base font-extrabold text-white disabled:cursor-not-allowed disabled:bg-[#e3ddd5] disabled:text-[#5f5851]"
              >
                {isSavingAll ? 'Saving…' : `Save all (${pendingSaves.size})`}
              </button>
            ) : null}
            <LibraryPager />
          </div>
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
    </LibraryPendingSavesContext>
  );
}
