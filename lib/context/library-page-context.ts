import { createContext, Dispatch, SetStateAction, use } from 'react';
import { LibraryGenreFilter, LibraryList } from '@/library/page';
import { MediaType } from 'lib/constants/media-types';
import { DatabaseSortOption } from 'lib/constants/database-sort-options';

export interface LibraryPageContextValue {
  page: number;
  setPage: Dispatch<SetStateAction<number>>;
  libraryList: LibraryList;
  type: MediaType;
  setType: Dispatch<SetStateAction<MediaType>>;
  limit: 12 | 24 | 48;
  setLimit: Dispatch<SetStateAction<12 | 24 | 48>>;
  sortBy: DatabaseSortOption;
  setSortBy: Dispatch<SetStateAction<DatabaseSortOption>>;
  genre: LibraryGenreFilter;
  setGenre: Dispatch<SetStateAction<LibraryGenreFilter>>;
  ascDesc: 'asc' | 'desc';
  setAscDesc: Dispatch<SetStateAction<'asc' | 'desc'>>;
  setTitleSearch: Dispatch<SetStateAction<string>>;
  refreshLibrary: () => Promise<void>;
}

const LibraryPageContext = createContext<LibraryPageContextValue | undefined>(
  undefined,
);

export function useLibraryPageContext() {
  const libraryPage = use(LibraryPageContext);
  if (!libraryPage)
    throw new Error('Must access library page context inside the provider');
  return libraryPage;
}

export default LibraryPageContext;
