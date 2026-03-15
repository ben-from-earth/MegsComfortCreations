'use client';

// react, redux imports
import { useEffect, useMemo, useRef, useState } from 'react';

// library imports
import { trpc } from 'lib/trpc/client';

// components
import DatabaseItemsContainer from '@/showdatabase/DatabaseItemsContainer';
import PaginationInputs from '@/showdatabase/PaginationInputs';

// context
import DatabasePageContext from 'lib/context/DatabasePageContext';

// helpers
import { titleRearrange } from 'lib/helpers/titleRearrange';

// interfaces and types
import {
  PostSavedMediaItem,
  SuccessfulPaginationResponse,
} from 'lib/interfaces/globalInterfaces';
import { MediaType } from 'lib/constants/mediaTypes';
import { DatabasePageContextValue } from 'lib/context/DatabasePageContext';
import { allGenres, NO_GENRE_FILTER } from '@/lib/enums/genreEnums';
import { DatabaseSortOption } from 'lib/constants/databaseSortOptions';

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
        // trigger refetches based on current inputs
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
