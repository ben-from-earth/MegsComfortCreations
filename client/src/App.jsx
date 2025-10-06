//package imports
import axios from 'axios';

//react router modules
import {
  createBrowserRouter,
  createRoutesFromElements,
  Route,
  RouterProvider,
} from 'react-router';

//layouts
import RootLayout from '@/layouts/RootLayout';

// pages
import ShowDatabasePage from '@/pages/ShowDatabase/ShowDatabasePage';
import MediaCollector from '@/pages/MediaCollector/MediaCollector';

//Context
import GenreContext from '@/context/GenreContext';
import { useEffect } from 'react';

//server location import from .env
const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

const router = createBrowserRouter(
  createRoutesFromElements(
    <Route path="/" element={<RootLayout />}>
      <Route path="ShowDatabase" element={<ShowDatabasePage />} />
      <Route index element={<MediaCollector />} />
    </Route>,
  ),
);

function App() {
  //get genres for use around the app
  const genres = [];

  useEffect(() => {
    (async () => {
      try {
        const res = await axios.get(`${serverDomain}/genres/getAll`);
        const collection = res.data;
        genres.push(...collection.genres);
      } catch (err) {
        console.error('Could not fetch genres: Server down or not active');
        return [];
      }
    })();
  }, []);

  return (
    <GenreContext.Provider value={genres}>
      <RouterProvider router={router} />
    </GenreContext.Provider>
  );
}

export default App;
