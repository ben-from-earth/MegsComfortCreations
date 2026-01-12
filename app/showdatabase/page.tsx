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
  SuccessfulMediaSearchResponse,
  SuccessfulPaginationResponse,
} from 'lib/interfaces/globalInterfaces';
import { DatabasePageContextValue } from 'lib/context/DatabasePageContext';
import { SortOptions } from '@/showdatabase/PaginationInputs';

export interface displayDatabaseItems {
  type: MediaType;
  items: PostSavedMediaItem[];
  total: number;
  min: number;
  max: number;
}

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
  const [genre, setGenre] = useState('');
  const [ascDesc, setAscDesc] = useState<'asc' | 'desc'>('asc');
  const effectiveSort = useMemo(() => {
    if (type === 'book') {
      if (sortBy === 'title') return 'title';
      if (sortBy === 'pubYear') return 'pubYear';
      return 'spineColor';
    }
    return 'title';
  }, [type, sortBy]);

  const searchQuery = trpc.database.searchByTitle.useQuery(
    { type, title: titleRearrange(titleSearch) },
    { enabled: titleSearch.length > 0 },
  );
  const paginatedQuery = trpc.database.getPaginated.useQuery(
    { type, sort: effectiveSort, limit, page, ascDesc },
    { enabled: titleSearch.length === 0 && (type !== 'book' || genre === '') },
  );
  const noGenresQuery = trpc.genres.paginateNoGenres.useQuery(
    { sort: effectiveSort, limit, page, ascDesc },
    {
      enabled: titleSearch.length === 0 && type === 'book' && genre === 'none',
    },
  );
  const byGenreQuery = trpc.genres.paginateByGenre.useQuery(
    { genre, sort: effectiveSort, limit, page, ascDesc },
    {
      enabled:
        titleSearch.length === 0 &&
        type === 'book' &&
        genre !== '' &&
        genre !== 'none',
    },
  );

  useEffect(() => {
    if (searchQuery.data) {
      const r = searchQuery.data as SuccessfulMediaSearchResponse;
      setDatabaseItems({
        type,
        items: r.foundMediaList,
        total: r.total,
        min: 1,
        max: r.total,
      });
    } else if (paginatedQuery.data) {
      const r = paginatedQuery.data as SuccessfulPaginationResponse;
      setDatabaseItems({
        type,
        items: r.paginatedList,
        total: r.total,
        min: (page - 1) * limit + 1,
        max: page * limit,
      });
    } else if (noGenresQuery.data) {
      const r = noGenresQuery.data as SuccessfulPaginationResponse;
      setDatabaseItems({
        type,
        items: r.paginatedList,
        total: r.total,
        min: (page - 1) * limit + 1,
        max: page * limit,
      });
    } else if (byGenreQuery.data) {
      const r = byGenreQuery.data as SuccessfulPaginationResponse;
      setDatabaseItems({
        type,
        items: r.paginatedList,
        total: r.total,
        min: (page - 1) * limit + 1,
        max: page * limit,
      });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    searchQuery.data,
    paginatedQuery.data,
    noGenresQuery.data,
    byGenreQuery.data,
  ]);

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
        searchQuery.refetch();
        paginatedQuery.refetch();
        noGenresQuery.refetch();
        byGenreQuery.refetch();
        return Promise.resolve();
      },
    }),
    [
      page,
      databaseItems,
      type,
      limit,
      sortBy,
      genre,
      ascDesc,
      searchQuery,
      paginatedQuery,
      noGenresQuery,
      byGenreQuery,
    ],
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
