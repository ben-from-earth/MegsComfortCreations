'use client';

import { useEffect, useMemo, useRef, useState } from 'react';
import { trpc } from 'lib/trpc/client';
import LibraryItems from '@/library/library-items';
import LibraryHeader from '@/library/library-header';
import LibraryPageContext, {
  type LibraryPageContextValue,
} from 'lib/context/library-page-context';
import { titleRearrange } from 'lib/helpers/title-rearrange';
import { PostSavedMediaItem } from 'lib/interfaces/global-interfaces';
import { MediaType } from 'lib/constants/media-types';
import { allGenres, NO_GENRE_FILTER } from '@/lib/enums/genre-enums';
import { DatabaseSortOption } from 'lib/constants/database-sort-options';

export interface LibraryList {
  type: MediaType;
  items: PostSavedMediaItem[];
  total: number;
  min: number;
  max: number;
}

export type LibraryGenreFilter =
  | (typeof allGenres)[number]
  | ''
  | typeof NO_GENRE_FILTER;

export default function LibraryPage() {
  const [libraryList, setLibraryList] = useState<LibraryList>({
    type: 'book',
    items: [],
    total: 0,
    min: 0,
    max: 0,
  });

  const [type, setType] = useState<MediaType>('book');
  const [limit, setLimit] = useState<12 | 24 | 48>(12);
  const [sortBy, setSortBy] = useState<DatabaseSortOption>('title');
  const [page, setPage] = useState<number>(1);
  const [titleSearch, setTitleSearch] = useState('');
  const [genre, setGenre] = useState<LibraryGenreFilter>('');
  const [ascDesc, setAscDesc] = useState<'asc' | 'desc'>('asc');
  const latestPaginationStateRef = useRef({ type, page, limit });
  const paginatedQuery = trpc.database.getPaginated.useQuery({
    type,
    sort: type === 'book' ? sortBy : 'title',
    limit,
    page,
    ascDesc,
    genre,
    title: titleRearrange(titleSearch),
  });

  useEffect(() => {
    latestPaginationStateRef.current = { type, page, limit };
  }, [type, page, limit]);

  useEffect(() => {
    if (!paginatedQuery.data) return;
    const { type: currentType, page: currentPage, limit: currentLimit } =
      latestPaginationStateRef.current;
    const pagination = paginatedQuery.data;
    setLibraryList({
      type: currentType,
      items: pagination.paginatedList,
      total: pagination.total,
      min: (currentPage - 1) * currentLimit + 1,
      max: Math.min(currentPage * currentLimit, pagination.total),
    });
  }, [paginatedQuery.data]);

  const libraryPageContextValue: LibraryPageContextValue = useMemo(
    () => ({
      page,
      setPage,
      libraryList,
      type,
      setType,
      limit,
      setLimit,
      sortBy,
      setSortBy,
      genre,
      setGenre,
      ascDesc,
      setAscDesc,
      setTitleSearch,
      refreshLibrary: async () => {
        await paginatedQuery.refetch();
      },
    }),
    [page, libraryList, type, limit, sortBy, genre, ascDesc, paginatedQuery],
  );

  return (
    <LibraryPageContext.Provider value={libraryPageContextValue}>
      <main className="mx-auto flex max-w-440 flex-col gap-4.5 px-6 pt-7 pb-16">
        <LibraryHeader />
        <LibraryItems />
      </main>
    </LibraryPageContext.Provider>
  );
}
