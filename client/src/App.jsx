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
import RootLayout from './layouts/RootLayout';

// pages
import HomePage from './pages/Home/HomePage';
import ShopPage from './pages/Shop/ShopPage';
import MegsRecs from './pages/MegsRecs/MegsRecs';
import NewsletterPage from './pages/Newsletter/NewsletterPage';
import ShowDatabasePage from './pages/ShowDatabase/ShowDatabasePage';
import MediaCollector from './pages/MediaCollector/MediaCollector';

//Context
import GenreContext from './context/GenreContext';

//server location import from .env
const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

//get genres for use around the app
const genres = await (async () => {
  try {
    const res = await axios.get(`${serverDomain}/genres/getAll`);
    const collection = res.data;
    return collection.genres;
  } catch (err) {
    console.error('Could not fetch genres: Server down or not active');
    return [];
  }
})();

const router = createBrowserRouter(
  createRoutesFromElements(
    <Route path="/" element={<RootLayout />}>
      <Route index element={<HomePage />} />
      <Route path="Shop" element={<ShopPage />} />
      <Route path="MegsRecs" element={<MegsRecs />} />
      <Route path="Newsletter" element={<NewsletterPage />} />
      <Route path="ShowDatabase" element={<ShowDatabasePage />} />
      <Route path="MediaCollector" element={<MediaCollector />} />
    </Route>,
  ),
);

function App() {
  return (
    <GenreContext.Provider value={genres}>
      <RouterProvider router={router} />
    </GenreContext.Provider>
  );
}

export default App;
