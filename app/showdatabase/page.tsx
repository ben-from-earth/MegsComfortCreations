'use client';

// react, redux imports
import { useCallback, useEffect, useMemo, useState } from 'react';

// library imports
import axios from 'axios';

// components
import DatabaseItemsContainer from '@/app/showdatabase/DatabaseItemsContainer';
import PaginationInputs from '@/app/showdatabase/PaginationInputs';

// context
import DatabasePageContext from '@/lib/context/DatabasePageContext';

// helpers
import { titleRearrange } from '@/lib/helpers/titleRearrange';

// interfaces and types
import {
  MediaType,
  PostSavedMediaItem,
  SuccessfulMediaSearchResponse,
  SuccessfulPaginationResponse,
} from '@/lib/interfaces/globalInterfaces';
import { DatabasePageContextValue } from '@/lib/context/DatabasePageContext';
import { SortOptions } from '@/app/showdatabase/PaginationInputs';

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

  const handleGetMedia = useCallback(async () => {
    if (titleSearch.length > 0) {
      try {
        const res = await axios.get<SuccessfulMediaSearchResponse>(
          `/api/database/search?type=${type}&title=${titleRearrange(titleSearch)}`,
        );
        const databaseResults = res.data;
        setDatabaseItems({
          type,
          items: databaseResults.foundMediaList,
          total: databaseResults.total,
          min: 1,
          max: databaseResults.total,
        });
      } catch (err) {
        console.log(`Server Error getting database items:`, err);
      }
    } else {
      if (type !== 'book' || genre === '') {
        try {
          const res = await axios.get<SuccessfulPaginationResponse>(
            `/api/database?type=${type}&sort=${sortBy}&limit=${limit}&page=${page}&ascDesc=${ascDesc}`,
          );
          const databaseResults = res.data;

          setDatabaseItems({
            type,
            items: databaseResults.paginatedList,
            total: databaseResults.total,
            min: (page - 1) * limit + 1,
            max: page * limit,
          });
        } catch (err) {
          console.log(`Server Error getting database items:`, err);
        }
      } else {
        if (genre === 'none') {
          const res = await axios.get<SuccessfulPaginationResponse>(
            `/api/genres/nogenres?sort=${sortBy}&limit=${limit}&page=${page}&ascDesc=${ascDesc}`,
          );
          const databaseResults = res.data;

          setDatabaseItems({
            type,
            items: databaseResults.paginatedList,
            total: databaseResults.total,
            min: (page - 1) * limit + 1,
            max: page * limit,
          });
        } else {
          const res = await axios.get<SuccessfulPaginationResponse>(
            `/api/genres?genre=${genre}&sort=${sortBy}&limit=${limit}&page=${page}&ascDesc=${ascDesc}`,
          );
          const databaseResults = res.data;

          setDatabaseItems({
            type,
            items: databaseResults.paginatedList,
            total: databaseResults.total,
            min: (page - 1) * limit + 1,
            max: page * limit,
          });
        }
      }
    }
  }, [page, type, limit, sortBy, titleSearch, genre, ascDesc]);

  useEffect(() => {
    handleGetMedia();
  }, [handleGetMedia]);

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
      handleGetMedia,
    }),
    [handleGetMedia, page, databaseItems, type, limit, sortBy, genre, ascDesc],
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
