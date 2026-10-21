'use client';

import { useEffect, useMemo, useRef, useState } from 'react';
import { trpc } from 'lib/trpc/client';
import DatabaseItemsContainer from '@/showdatabase/database-items-container';
import PaginationInputs from '@/showdatabase/pagination-inputs';
import DatabasePageContext, {
  type DatabasePageContextValue,
} from 'lib/context/database-page-context';
import { titleRearrange } from 'lib/helpers/title-rearrange';
import {
  PostSavedMediaItem,
  SuccessfulPaginationResponse,
} from 'lib/interfaces/global-interfaces';
import { MediaType } from 'lib/constants/media-types';
import { allGenres, NO_GENRE_FILTER } from '@/lib/enums/genre-enums';
import { DatabaseSortOption } from 'lib/constants/database-sort-options';

export interface displayDatabaseItems {
  type: MediaType;
  items: PostSavedMediaItem[];
  total: number;
  min: number;
  max: number;
}

export type genreInput =
  | (typeof allGenres)[number]
  | ''
  | typeof NO_GENRE_FILTER;

export default function ShowDatabase() {
  const [databaseItems, setDatabaseItems] = useState<displayDatabaseItems>({
    type: 'book',
    items: [],
    total: 0,
    min: 0,
    max: 0,
  });

  const [type, setType] = useState<MediaType>('book');
  const [limit, setLimit] = useState<3 | 5 | 10>(5);
  const [sortBy, setSortBy] = useState<DatabaseSortOption>('title');
  const [page, setPage] = useState<number>(1);
  const [titleSearch, setTitleSearch] = useState('');
  const [genre, setGenre] = useState<genreInput>('');
  const [ascDesc, setAscDesc] = useState<'asc' | 'desc'>('asc');
  const latestPaginationStateRef = useRef({ type, page, limit });
  const effectiveSort = useMemo(() => {
    if (type === 'book') {
      return sortBy;
    }
    return 'title';
  }, [type, sortBy]);

  const paginatedQuery = trpc.database.getPaginated.useQuery({
    type,
    sort: effectiveSort,
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
    const r = paginatedQuery.data as SuccessfulPaginationResponse;
    setDatabaseItems({
      type: currentType,
      items: r.paginatedList,
      total: r.total,
      min: (currentPage - 1) * currentLimit + 1,
      max: Math.min(currentPage * currentLimit, r.total),
    });
  }, [paginatedQuery.data]);

  const DatabasePageContextValue: DatabasePageContextValue = useMemo(
    () => ({
      page,
      setPage,
      databaseItems,
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
      handleGetMedia: async () => {
        await paginatedQuery.refetch();
      },
    }),
    [page, databaseItems, type, limit, sortBy, genre, ascDesc, paginatedQuery],
  );

  return (
    <DatabasePageContext.Provider value={DatabasePageContextValue}>
      <div className="flex flex-col items-center">
        <PaginationInputs />
        <DatabaseItemsContainer />
      </div>
    </DatabasePageContext.Provider>
  );
}
