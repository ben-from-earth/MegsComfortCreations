'use client';

// react, redux imports
import { useEffect, useMemo, useState } from 'react';

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
  MediaType,
  PostSavedMediaItem,
  SuccessfulPaginationResponse,
} from 'lib/interfaces/globalInterfaces';
import { DatabasePageContextValue } from 'lib/context/DatabasePageContext';
import { SortOptions } from '@/showdatabase/PaginationInputs';
import { allGenres } from '@/lib/enums/genreEnums';

export interface displayDatabaseItems {
  type: MediaType;
  items: PostSavedMediaItem[];
  total: number;
  min: number;
  max: number;
}

export type genreInput = (typeof allGenres)[number] | '' | 'None';

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
  const [sortBy, setSortBy] = useState<SortOptions>('title');
  const [page, setPage] = useState<number>(1);
  const [titleSearch, setTitleSearch] = useState('');
  const [genre, setGenre] = useState<genreInput>('');
  const [ascDesc, setAscDesc] = useState<'asc' | 'desc'>('asc');
  const effectiveSort = useMemo(() => {
    if (type === 'book') {
      if (sortBy === 'title') return 'title';
      if (sortBy === 'pubYear') return 'pubYear';
      return 'spineColor';
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
    if (!paginatedQuery.data) return;
    const r = paginatedQuery.data as SuccessfulPaginationResponse;
    setDatabaseItems({
      type,
      items: r.paginatedList,
      total: r.total,
      min: (page - 1) * limit + 1,
      max: Math.min(page * limit, r.total),
    });
  }, [paginatedQuery.data, page, limit, type]);

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
