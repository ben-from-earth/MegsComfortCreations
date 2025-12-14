// react, redux imports
import {
  createAsyncThunk,
  createSlice,
  nanoid,
  PayloadAction,
} from '@reduxjs/toolkit';

// library imports
import axios from 'axios';

// helpers
import { titleRearrange } from '@/lib/helpers/titleRearrange';
import { updateQueryCount } from '@/lib/helpers/updateQueryCount';

// interfaces and types
import { RootState } from '@/lib/state/store';
import {
  BlockInfo,
  BookRow,
  MediaLabel,
  MediaType,
  MovieRow,
  SuccessfulMediaSearchResponse,
} from '@/lib/interfaces/globalInterfaces';
import { OpenLibraryError, SearchErrorResponse } from '@/app/api/api-Errors';
import { OpenLibrarySuccess } from '@/app/api/getonlinedata/openlibrary/route';
import { titleOutputObj } from '@/lib/helpers/titleCollectionListConversion';

export interface mediaTypeDefinitions {
  type: MediaType;
  label: MediaLabel;
  show: boolean;
  titleCollectionList: titleOutputObj[];
}

interface InitialState {
  mediaTypeDefinitions: mediaTypeDefinitions[];
  shouldFetch: boolean;
  isLoading: boolean;
}

// shared fields for all media
type BaseBlockInfo = {
  title: string;
  spineColor?: string;
  databaseGenres?: string[]; // only really used for books, but harmless here
};

// extra fields only for books
type BookBlockInfo = BaseBlockInfo & {
  author?: string;
  pubYear?: number | null;
  pageCount?: number | null;
};

// discriminated union for the whole block
export type CollectedBlockInformation =
  | {
      type: 'book';
      images: string[];
      BlockInfo: BookBlockInfo;
      blockID: string;
      isDatabase: boolean;
    }
  | {
      type: 'movie' | 'video_game' | 'album';
      images: string[];
      BlockInfo: BaseBlockInfo;
      blockID: string;
      isDatabase: boolean;
    };

//set up media types and respective labels
export const medias: {
  type: MediaType;
  label: MediaLabel;
}[] = [
  { type: 'book', label: 'Book' },
  { type: 'movie', label: 'Movie' },
  { type: 'video_game', label: 'Video Game' },
  { type: 'album', label: 'Album' },
];

// state holds booleans for should fetch and loading, and an array of the media types
// mediaType: {type, label, show (for checkboxes), and titleCollectionList (data used to fetch Google Search)}
const initialState: InitialState = {
  mediaTypeDefinitions: medias.map(({ type, label }) => ({
    type,
    label,
    show: false,
    titleCollectionList: [],
  })),
  shouldFetch: false,
  isLoading: false,
};

export const collectBlockInformation = createAsyncThunk(
  'collector/getMediaCovers',
  async ({
    type,
    toCollectItem,
  }: {
    type: MediaType;
    toCollectItem: titleOutputObj;
  }) => {
    const { title, author } = toCollectItem;

    //check database for existing data with same title.
    const mediaSearchRes = await axios.get(`/api/database/search`, {
      params: { type, title: titleRearrange(title) },
      //accept 400 codes for error handling
      validateStatus: (status) => status < 500,
    });
    const mediaSearchData: SearchErrorResponse | SuccessfulMediaSearchResponse =
      mediaSearchRes.data;

    //if we return a book from the database, return the information.
    if ('foundMediaList' in mediaSearchData) {
      //--- still need to write logic for more than one return ---//
      //still checking only first index here

      if (type === 'book') {
        const { id, imageUrls, title, author, pageCount, pubYear, spineColor } =
          mediaSearchData.foundMediaList[0] as BookRow;
        //get genres tied to the found book id
        const genreSearchRes = await axios.get(
          `/api/genres/getforbook?bookID=${id}`,
        );
        const genreSearchData: { message: string; genres: string[] } =
          genreSearchRes.data;
        const databaseGenres = genreSearchData.genres;

        //return all the block info and designate isDatabase to be true for books
        return {
          type,
          images: imageUrls,
          BlockInfo: {
            title,
            author,
            pubYear,
            pageCount,
            spineColor,
            databaseGenres,
          },
          blockID: nanoid(),
          isDatabase: true,
        };
      }

      const { imageUrls, title, spineColor } = mediaSearchData
        .foundMediaList[0] as MovieRow;

      //return all the block info and designate isDatabase to be true for other media
      return {
        type,
        images: imageUrls,
        BlockInfo: {
          title,
          spineColor,
        },
        blockID: nanoid(),
        isDatabase: true,
      };
    }

    //if the media wasnt in the database, collect cover images
    const imageSearchRes = await axios.post<{ images: string[] }>(
      `/api/getonlinedata/mediacovers`,
      {
        title,
        author,
        type,
      },
    );
    //conservative query count update every time we make a request to google search API.
    updateQueryCount();
    const imgArr = imageSearchRes.data.images;

    let BlockInfo: BlockInfo;
    if (type === 'book') {
      //if book, go to open library and get more data about the book
      const openLibraryRes = await axios.post<
        OpenLibrarySuccess | OpenLibraryError
      >(`/api/getonlinedata/openlibrary`, { title, author });

      const bookInformation = openLibraryRes.data;

      if ('failedSearchData' in bookInformation) {
        BlockInfo = bookInformation.failedSearchData;
      } else {
        BlockInfo = bookInformation;
      }
      // BlockInfo: { title, author, pubYear, pageCount } || {title, author}
    } else {
      //Just submit title as BlockInfo for non-books
      //Updates to data collection for other media types can be performed here if necessary in future update.
      BlockInfo = { title };
    }

    const CollectedBlockInformation: CollectedBlockInformation = {
      type,
      images: imgArr,
      BlockInfo,
      blockID: nanoid(),
      isDatabase: false,
    };

    //return the collected data for creation of collectedCoverBlock
    return CollectedBlockInformation;
  },
);

export const collectorSlice = createSlice({
  name: 'collector',
  initialState,
  reducers: {
    //function to handle showing the media collector text area if the checkbox is selected
    setChecks: (state, action): void => {
      const idx: number = action.payload;
      state.mediaTypeDefinitions[idx].show =
        !state.mediaTypeDefinitions[idx].show;
    },

    // takes in the text area text and creates a list of search items.
    // books are inputted as title / author, title / author, etc. so we parse out the string here
    setCollectList: (
      state,
      action: PayloadAction<
        { type: MediaType; titleSearchList: titleOutputObj[] }[]
      >,
    ): void => {
      for (const media of action.payload) {
        const i: number = state.mediaTypeDefinitions.findIndex(
          (m) => m.type === media.type,
        );
        if (i !== -1)
          state.mediaTypeDefinitions[i].titleCollectionList =
            media.titleSearchList;
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
      state.mediaTypeDefinitions = state.mediaTypeDefinitions.map((mt) => ({
        ...mt,
        titleCollectionList: [],
      }));
    },
  },
});

export const mediaData = (state: RootState): mediaTypeDefinitions[] =>
  state.collector.mediaTypeDefinitions;
export const getFetchStatus = (state: RootState): boolean =>
  state.collector.shouldFetch;
export const getLoadingStatus = (state: RootState): boolean =>
  state.collector.isLoading;
export const {
  setChecks,
  setCollectList,
  startLoad,
  startFetch,
  finishedLoad,
  finishedFetch,
} = collectorSlice.actions;
export default collectorSlice.reducer;
