//package imports
import axios from 'axios';
import { useEffect, useState } from 'react';

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
import HomePage from '@/pages/HomePage/HomePage';
import LoginPage from '@/pages/LoginPage/LoginPage';
import SignupPage from '@/pages/SignupPage/SignupPage';
import ErrorBoundary from '@/pages/ErrorBoundary/ErrorBoundary';

//Context
import GenreContext from '@/context/GenreContext';
import ProfilePage from '@/pages/ProfilePage/ProfilePage';
import requireAuth from '@/pages/MediaCollector/helpers/authCheck';

//server location import from .env
const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

const router = createBrowserRouter(
  createRoutesFromElements(
    <Route path="/" element={<RootLayout />} errorElement={<ErrorBoundary />}>
      <Route index element={<HomePage />} />
      <Route
        path="ShowDatabase"
        element={<ShowDatabasePage />}
        loader={requireAuth}
      />
      <Route path="login" element={<LoginPage />} />
      <Route path="signup" element={<SignupPage />} />
      <Route
        path="MediaCollector"
        element={<MediaCollector />}
        loader={requireAuth}
      />
      <Route path="Profile" element={<ProfilePage />} loader={requireAuth} />
    </Route>,
  ),
);

function App() {
  //get genres for use around the app
  const [genres, setGenres] = useState([]);

  useEffect(() => {
    (async () => {
      try {
        const res = await axios.get(`${serverDomain}/genres/getAll`);
        const collection = res.data;
        setGenres(collection.genres);
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
