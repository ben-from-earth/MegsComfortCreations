'use client';

// react imports
import { useEffect, useState } from 'react';

// context
import GenreContext from 'lib/context/GenreContext';

// interfaces and types
import { trpc } from 'lib/trpc/client';
import { Genre } from './lib/enums/genreEnums';

export default function Providers({ children }: { children: React.ReactNode }) {
  //get genres for use around the app
  const [genres, setGenres] = useState<Genre[]>([]);
  const { data } = trpc.genres.getAll.useQuery(undefined, {
    staleTime: 5 * 60 * 1000,
  });
  useEffect(() => {
    if (data?.genres) setGenres(data.genres);
  }, [data]);
  return (
    <GenreContext.Provider value={genres}>{children}</GenreContext.Provider>
  );
}
