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
  return <RouterProvider router={router} />;
}

export default App;
