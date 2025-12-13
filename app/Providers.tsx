'use client';

// react, redux imports
import { Provider } from 'react-redux';
import { store } from '@/lib/state/store';
import { useEffect, useState } from 'react';

// library imports
import axios from 'axios';

// context
import GenreContext from '@/lib/context/GenreContext';

// interfaces and types
import { getAllResponse } from '@/app/api/genres/getall/route';

export default function Providers({ children }: { children: React.ReactNode }) {
  //get genres for use around the app
  const [genres, setGenres] = useState<string[]>([]);

  useEffect(() => {
    (async () => {
      try {
        const res = await axios.get<getAllResponse>(`/api/genres/getall`);
        const collection = res.data;
        setGenres(collection.genres);
      } catch {
        console.error('Could not fetch genres: Server down or not active');
        return [];
      }
    })();
  }, []);
  return (
    <Provider store={store}>
      <GenreContext.Provider value={genres}>{children}</GenreContext.Provider>
    </Provider>
  );
}
