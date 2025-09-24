//redux
import { createAsyncThunk, createSlice, nanoid } from '@reduxjs/toolkit';

//axios
import axios from 'axios';

//helpers
import {
  titleRearrange,
  updateQueryCount,
} from '@/pages/MediaCollector/helpers/mediaCollectorHelpers';

//server domain for axios requests
const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

//set up media types and respective labels
export const medias = [
  { type: 'book', label: 'Book' },
  { type: 'movie', label: 'Movie' },
  { type: 'video_game', label: 'Video Game' },
  { type: 'album', label: 'Album' },
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

export const collectBlockInformation = createAsyncThunk(
  'collector/getMediaCovers',
  async ({ type, toCollectItem }) => {
    //setup search inputs based on media type
    let title;
    let author;
    if (type === 'book') {
      title = toCollectItem.title;
      author = toCollectItem.author;
    } else {
      title = toCollectItem;
    }

    //check database for existing data with same title.
    const mediaSearchRes = await axios.get(`${serverDomain}/database/search`, {
      params: { type, title: titleRearrange(title) },
      //accept 400 codes for error handling
      validateStatus: (status) => status < 500,
    });
    const mediaSearchData = mediaSearchRes.data;

    //if we return a book from the database, return the information.
    if (mediaSearchData.foundMediaList?.length > 0) {
      //--- still need to write logic for more than one return ---//
      const {
        id,
        image_urls,
        title,
        author,
        page_count,
        pub_year,
        spine_color,
      } = mediaSearchData.foundMediaList[0]; //still checking only first index here

      if (type === 'book') {
        //get genres tied to the found book id
        const genreSearchRes = await axios.post(
          `${serverDomain}/genres/getForBook`,
          { bookID: id },
        );
        const genreSearchData = genreSearchRes.data;
        const databaseGenres = genreSearchData.genres;

        //return all the block info and designate isDatabase to be true
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

      //return all the block info and designate isDatabase to be true
      return {
        type,
        images: image_urls,
        blockInfo: {
          title,
          spine_color,
        },
        blockID: nanoid(),
        isDatabase: true,
      };
    }

    //if media wasnt in database, collect cover images:

    const imageSearchRes = await axios.post(
      `${serverDomain}/getOnlineData/mediacovers`,
      {
        title,
        author,
        type,
      },
    );
    //conservative query count update every time we make a request to google search API.
    updateQueryCount();
    const imgArr = imageSearchRes.data.images;

    let blockInfo;
    if (type === 'book') {
      //if book, go to open library and get more data about the book
      const openLibraryRes = await axios.post(
        `${serverDomain}/getOnlineData/openlibrary`,
        { title, author },
      );
      if (openLibraryRes.data.errors) {
        blockInfo = { title, author };
      } else {
        blockInfo = openLibraryRes.data;
      }
      // blockInfo: { title, author, pub_year, page_count } || {title, author}
    } else {
      //Just submit title as blockInfo for non-books
      //Updates to data collection for other media types can be performed here if necessary in future update.
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
  },
);

export const collectorSlice = createSlice({
  name: 'collector',
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
        if (media.type === 'book' && media.text) {
          searchArr = media.text
            .split(',')
            .map((i) => i.trim())
            .filter((i) => i !== '');
          searchArr = searchArr.map((t) => {
            const titleInfo = t.split('/').map((i) => i.trim());
            const title = titleInfo[0];
            const author = titleInfo[1];

            return {
              title,
              author,
            };
          });
        } else if (media.text) {
          searchArr = media.text
            .split(',')
            .map((i) => i.trim())
            .filter((i) => i !== '');
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
