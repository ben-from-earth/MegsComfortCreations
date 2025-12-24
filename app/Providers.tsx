'use client';

// react, redux imports
import { Provider } from 'react-redux';
import { store } from '@/lib/state/store';
import { useEffect, useState } from 'react';

// library imports
// axios no longer needed for genres; using tRPC

// context
import GenreContext from '@/lib/context/GenreContext';

// interfaces and types
import { trpc } from '@/lib/trpc/client';

export default function Providers({ children }: { children: React.ReactNode }) {
  //get genres for use around the app
  const [genres, setGenres] = useState<string[]>([]);
  const { data } = trpc.genres.getAll.useQuery(undefined, {
    staleTime: 5 * 60 * 1000,
  });
  useEffect(() => {
    if (data?.genres) setGenres(data.genres);
  }, [data]);
  return (
    <Provider store={store}>
      <GenreContext.Provider value={genres}>{children}</GenreContext.Provider>
    </Provider>
  );
}
