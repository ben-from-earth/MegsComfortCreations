import "./App.css";
import {
  createBrowserRouter,
  createRoutesFromElements,
  Route,
  RouterProvider,
} from "react-router";

//layouts
import RootLayout from "./layouts/RootLayout";

// pages
import HomePage from "./pages/Home/HomePage";
import ShopPage from "./pages/Shop/ShopPage";
import MegsRecs from "./pages/MegsRecs/MegsRecs";
import NewsletterPage from "./pages/Newsletter/NewsletterPage";
import MediaCollector from "./pages/MediaCollector/MediaCollector";

//Context
import GenreContext from "./context/GenreContext";

//get genres for use around the app
const genres = await (async () => {
  try {
    const res = await fetch("http://localhost:3001/genres/getAll");

    if (!res.ok) {
      throw new Error(`Server Error getting genres: ${res.status}`);
    }
    const collection = await res.json();
    return collection.payload?.map((item) => item.genre);
  } catch (err) {
    console.error("Could not fetch genres: Server down or not active");
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
      <Route path="MediaCollector" element={<MediaCollector />} />
    </Route>
  )
);

function App() {
  return (
    <GenreContext.Provider value={genres}>
      <RouterProvider router={router} />
    </GenreContext.Provider>
  );
}

export default App;
