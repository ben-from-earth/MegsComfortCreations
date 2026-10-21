// react, redux imports
import { createContext, Dispatch, SetStateAction, useContext } from 'react';

// interfaces and types
import { displayDatabaseItems, genreInput } from '@/showdatabase/page';
import { MediaType } from 'lib/constants/media-types';
import { DatabaseSortOption } from 'lib/constants/database-sort-options';

export interface DatabasePageContextValue {
  page: number;
  setPage: Dispatch<SetStateAction<number>>;
  databaseItems: displayDatabaseItems;
  type: MediaType;
  setType: Dispatch<SetStateAction<MediaType>>;
  limit: number;
  setLimit: Dispatch<SetStateAction<3 | 5 | 10>>;
  sortBy: DatabaseSortOption;
  setSortBy: Dispatch<SetStateAction<DatabaseSortOption>>;
  genre: genreInput;
  setGenre: Dispatch<SetStateAction<genreInput>>;
  ascDesc: 'asc' | 'desc';
  setAscDesc: Dispatch<SetStateAction<'asc' | 'desc'>>;
  setTitleSearch: Dispatch<SetStateAction<string>>;
  handleGetMedia: () => Promise<void>;
}

const DatabasePageContext = createContext<DatabasePageContextValue | undefined>(
  undefined,
);

export function useDatabasePageContext() {
  const ctx = useContext(DatabasePageContext);
  if (!ctx)
    throw new Error('Must access database page context inside the provider');
  return ctx;
}

export default DatabasePageContext;
