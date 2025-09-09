import { createAsyncThunk, createSlice, nanoid } from "@reduxjs/toolkit";
import { updateQueryCount } from "../pages/MediaCollector/helpers/MediaCollectorHelpers";

//get Google Search API information from .env
const API_KEY = import.meta.env.VITE_GOOGLE_SEARCH_API_KEY;
const CX = import.meta.env.VITE_GOOGLE_SEARCH_CX;

//set up media types and respective labels
export const medias = [
  { type: "book", label: "Book" },
  { type: "movie", label: "Movie" },
  { type: "video_game", label: "Video Game" },
  { type: "album", label: "Album" },
];

// state holds booleans for should fetch and loading, and an array of the media types
// mediaType: {type, label, show (for checkboxes), and toCollect (data used to fetch Google Search)}
const initialState = {
  mediaTypes: medias.map(({ type, label }) => ({
    type,
    label,
    show: false,
    toCollect: [],
  })),
  shouldFetch: false,
  isLoading: false,
};

export const grabOpenLibraryData = createAsyncThunk(
  "collector/fetchLibraryData",
  async ({ title, author }) => {
    try {
      const params = new URLSearchParams({
        title,
        author,
        limit: "1",
        fields: "first_publish_year,number_of_pages_median",
      });
      const res = await fetch(
        `https://openlibrary.org/search.json?${params.toString()}`
      );
      if (!res.ok) {
        return { title, author };
      }

      const data = await res.json();
      const doc = data?.docs?.[0];
      if (!doc) {
        return { title, author };
      }

      const {
        first_publish_year: pub_year,
        number_of_pages_median: page_count,
      } = doc;
      return { title, author, pub_year, page_count };
    } catch {
      console.log(`Error gathering Open Library data for ${title}`);
      return { title, author };
    }
  }
);

const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

export const collectBlockInformation = createAsyncThunk(
  "collector/getMediaCovers",
  async ({ type, toCollectItem }, { signal, dispatch }) => {
    //setup search inputs based on media type
    let title;
    let author;
    if (type === "book") {
      title = toCollectItem.title;
      author = toCollectItem.author;
    } else {
      title = toCollectItem;
    }

    //check database for existing data with same title.
    const bookSearchRes = await fetch(
      `${serverDomain}/database/search?type=${type}&title=${title}`
    );
    const bookSearchData = await bookSearchRes.json();
    if (bookSearchData.foundBooksList?.length > 0) {
      //--- still need to write logic for more than one return ---//
      const {
        id,
        image_urls,
        title,
        author,
        page_count,
        pub_year,
        spine_color,
      } = bookSearchData.foundBooksList[0]; //still checking only first index here
      const genreSearchRes = await fetch(`${serverDomain}/genres/getFromBook`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          bookID: id,
        }),
      });
      const genreSearchData = await genreSearchRes.json();
      const databaseGenres = genreSearchData.payload;
      return {
        type,
        images: image_urls,
        blockInfo: {
          title,
          author,
          pub_year,
          page_count,
          spine_color,
          databaseGenres,
        },
        blockID: nanoid(),
        isDatabase: true,
      };
    }

    //setup empty img array to fill with searched images
    const imgArr = [];

    try {
      //search query example: "Dune book cover image"
      //request for three images from this search
      const params = new URLSearchParams({
        q: `${title} ${type} Cover Image`,
        cx: CX,
        key: API_KEY,
        searchType: "image",
        num: 3,
      });

      const res = await fetch(
        `https://www.googleapis.com/customsearch/v1?${params.toString()}`,
        { signal }
      );

      //this counts as a query even if failed so let the UI know
      updateQueryCount();

      //if some kind of API failure, deal with it here
      if (!res.ok) {
        throw new Error(
          `Google Search API failed: ${res.status} ${res.statusText}`
        );
      }

      //push the images into the array that was established earlier
      const imageURLs = await res.json();
      imageURLs.items.map((i) => imgArr.push(i.link));
    } catch (e) {
      if (e.name !== "AbortError") {
        console.log("Media Cover Collection error:", e.message);
      } else {
        console.log("Aborted");
      }
    }

    let blockInfo;
    if (type === "book") {
      try {
        //if book, go to open library and get more data about the book
        blockInfo = await dispatch(
          grabOpenLibraryData({ title, author })
        ).unwrap();
        // blockInfo: { title, author, pub_year, page_count } || {title, author}
      } catch (err) {
        console.log("Dispatch issue:", err);
      }
    } else {
      //Just submit title as blockInfo for non-books
      //Updates to data collection for other media types can be performed here
      blockInfo = { title };
    }
    //return the collected data for creation of collectedCoverBlock
    return {
      type,
      images: imgArr,
      blockInfo,
      blockID: nanoid(),
      isDatabase: false,
    };
  }
);

export const collectorSlice = createSlice({
  name: "collector",
  initialState,
  reducers: {
    //function to handle showing the media collector text area if the checkbox is selected
    setChecks: (state, action) => {
      const idx = action.payload;
      state.mediaTypes[idx].show = !state.mediaTypes[idx].show;
    },

    // takes in the text area text and creates a list of search items.
    // books are inputted as title / author, title / author, etc. so we parse out the string here
    setCollectText: (state, action) => {
      for (let media of action.payload.searchData) {
        let searchArr = [];
        if (media.type === "book" && media.text) {
          searchArr = media.text
            .split(",")
            .map((i) => i.trim())
            .filter((i) => i !== "");
          searchArr = searchArr.map((t) => {
            const titleInfo = t.split("/").map((i) => i.trim());
            const title = titleInfo[0];
            const author = titleInfo[1];

            return {
              title,
              author,
            };
          });
        } else if (media.text) {
          searchArr = media.text
            .split(",")
            .map((i) => i.trim())
            .filter((i) => i !== "");
        }
        const i = state.mediaTypes.findIndex((m) => m.type === media.type);
        if (i !== -1) state.mediaTypes[i].toCollect = searchArr;
      }
    },
    startLoad: (state) => {
      state.isLoading = true;
    },
    startFetch: (state) => {
      state.shouldFetch = true;
    },
    finishedLoad: (state) => {
      state.isLoading = false;
    },
    finishedFetch: (state) => {
      state.isLoading = false;
      state.shouldFetch = false;
      state.mediaTypes = state.mediaTypes.map((mt) => ({
        ...mt,
        toCollect: [],
      }));
    },
  },
});

export const mediaData = (state) => state.collector.mediaTypes;
export const getFetchStatus = (state) => state.collector.shouldFetch;
export const getLoadingStatus = (state) => state.collector.isLoading;
export const {
  setChecks,
  setCollectText,
  startLoad,
  startFetch,
  finishedLoad,
  finishedFetch,
} = collectorSlice.actions;
export default collectorSlice.reducer;
