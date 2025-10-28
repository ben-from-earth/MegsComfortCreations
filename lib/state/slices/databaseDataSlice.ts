// react, redux imports
import { createAsyncThunk, createSlice, PayloadAction } from '@reduxjs/toolkit';
import { RootState } from '@/lib/state/store';

// library imports
import axios from 'axios';

// necessary imports from collector state
import { medias } from '@/lib/state/slices/collectorSlice';

// helpers
import { titleRearrange } from '@/lib/helpers/titleRearrange';

// interfaces and types
import { databasePayload } from '@/app/mediacollector/CollectedCoversBlock';
import {
  databaseSaveServerResponse,
  MediaLabel,
  MediaType,
  presavedMediaItem,
  SuccessfulGenreLinkUnlinkResponse,
  SuccessfulMediaSaveEditResponse,
} from '@/lib/interfaces/globalInterfaces';
import { DatabaseSaveEditErrorResponse } from '@/app/api/api-Errors';

export interface databaseSliceData extends presavedMediaItem {
  images: { src: string; idx: number }[];
}
export interface databaseDataPerType {
  type: MediaType;
  label: MediaLabel;
  data: databaseSliceData[];
}

//set up initial state
const initialState: databaseDataPerType[] = medias.map(({ type, label }) => ({
  type,
  label,
  data: [] as databaseSliceData[],
}));

export const sendToDatabase = createAsyncThunk(
  '/databaseData/sendtodatabase',
  async (
    databaseData: databaseDataPerType[],
  ): Promise<databaseSaveServerResponse> => {
    const serverResponses: databaseSaveServerResponse = [];
    for (const media of databaseData) {
      const sendData = media.data;
      if (media.type === 'book') {
        const bookPromises = sendData.map(async (book: databaseSliceData) => {
          const title = titleRearrange(book.title);
          const bookData: presavedMediaItem = {
            title,
            author: book.author,
            page_count: book.page_count,
            pub_year: book.pub_year,
            image_urls: book.images.map((item) => item.src),
            spine_color: book.spine_color,
          };

          try {
            const bookSaveRes = await axios.post<
              SuccessfulMediaSaveEditResponse | DatabaseSaveEditErrorResponse
            >('api/database/save/book', bookData, {
              validateStatus: (status) => status < 500,
            });
            const bookCreationResponse = bookSaveRes.data;
            if ('error' in bookCreationResponse) {
              return bookCreationResponse;
            } else {
              const bookDatabaseID = bookCreationResponse.actionAttemptItem.id;

              const genreLinkRes = await axios.post<{
                genreResponses: SuccessfulGenreLinkUnlinkResponse[];
              }>(
                'api/genres/addlink',
                { bookID: bookDatabaseID, genres: book.genres },
                { validateStatus: (status) => status < 500 },
              );
              const genreLinkResponse = genreLinkRes.data;
              return { ...bookCreationResponse, ...genreLinkResponse };
            }
          } catch (error) {
            const serverError: DatabaseSaveEditErrorResponse = {
              actionAttemptItem: bookData,
              type: media.type,
              errors: [
                'Server Error during save',
                `${bookData.title} did not save to the database`,
              ],
              error: 'Server Error',
              message: `There was a server error during save attempt for ${bookData.title}`,
            };
            return serverError;
          }
        });
        const results = await Promise.allSettled(bookPromises);
        serverResponses.push(
          ...results
            .filter((r) => r.status === 'fulfilled')
            .map((result) => result.value),
        );
      } else {
        const otherMediaPromises = sendData.map(
          async (otherMedia: databaseSliceData) => {
            const title = titleRearrange(otherMedia.title);
            const otherMediaData: presavedMediaItem = {
              title,
              image_urls: otherMedia.images.map((img) => img.src),
              spine_color: otherMedia.spine_color,
            };

            try {
              const otherMediaSaveRes = await axios.post<
                SuccessfulMediaSaveEditResponse | DatabaseSaveEditErrorResponse
              >(`api/database/save/${media.type}`, otherMediaData, {
                validateStatus: (status) => status < 500,
              });
              const otherMediaCreationResponse = otherMediaSaveRes.data;
              return otherMediaCreationResponse;
            } catch (error) {
              const serverError: DatabaseSaveEditErrorResponse = {
                actionAttemptItem: otherMediaData,
                type: media.type,
                errors: [
                  'Server Error during save',
                  `${otherMediaData.title} did not save to the database`,
                ],
                error: 'Server Error',
                message: `There was a server error during save attempt for ${otherMediaData.title}`,
              };
              return serverError;
            }
          },
        );
        const results = await Promise.allSettled(otherMediaPromises);
        serverResponses.push(
          ...results
            .filter((r) => r.status === 'fulfilled')
            .map((result) => result.value),
        );
      }
    }

    return serverResponses;
  },
);

export const databaseDataSlice = createSlice({
  name: 'databaseData',
  initialState,
  reducers: {
    populateDatabaseData: (
      state,
      action: PayloadAction<databasePayload>,
    ): void => {
      const { type, data } = action.payload;
      const i: number = state.findIndex((m) => m.type === type);

      //check if the data already exists before pushing
      let exists = false;
      for (const item of state[i].data) {
        if (
          type === 'book' &&
          item.title === data.title &&
          item.author === data.author
        ) {
          exists = true;
        } else if (type !== 'book' && item.title === data.title) {
          exists = true;
        }
      }
      if (!exists) {
        state[i].data.push({ ...data, images: [] });
      }
    },
    clearDatabaseData: () => initialState,

    // if the user updates the data in the collectedCoverBlock text area,
    //reflect the changes in the state as well
    updateDatabaseData: (
      state,
      action: PayloadAction<{
        blockID: string;
        type: 'book' | 'movie' | 'video_game' | 'album';
        name: 'title' | 'author' | 'pub_year' | 'page_count';
        newText: string;
      }>,
    ) => {
      const { blockID, type, name, newText } = action.payload;
      const i = state.findIndex((m) => m.type === type);
      const j = state[i].data.findIndex((block) => block.blockID === blockID);
      if (name === 'pub_year' || name === 'page_count') {
        state[i].data[j][name] = Number(newText);
      } else {
        state[i].data[j][name] = newText;
      }
    },
    //push image to state storage if user selects it
    addToDatabaseData: (
      state,
      action: PayloadAction<{
        blockID: string;
        type: MediaType;
        idx?: number;
        src?: string;
        spine_color?: string;
        genreText?: string;
      }>,
    ) => {
      const { blockID, type, idx, src, spine_color, genreText } =
        action.payload;
      const i = state.findIndex((m) => m.type === type);
      const j = state[i].data.findIndex((block) => block.blockID === blockID);
      const chosenBlock = state[i].data[j];
      if (src) chosenBlock.images.push({ src, idx: idx! });
      if (spine_color) chosenBlock.spine_color = spine_color;
      if (genreText) chosenBlock.genres!.push(genreText);
    },

    //remove image from state storage if user deselects it
    removeFromDatabaseData: (
      state,
      action: PayloadAction<{
        blockID: string;
        type: MediaType;
        idx?: number;
        genreText?: string;
        deleteBlock?: boolean;
      }>,
    ) => {
      const { blockID, type, idx, genreText, deleteBlock } = action.payload;
      const i = state.findIndex((m) => m.type === type);
      const j = state[i].data.findIndex((block) => block.blockID === blockID);
      const chosenBlock = state[i].data[j];
      if (idx && isFinite(idx)) {
        chosenBlock.images = chosenBlock.images.filter(
          (img) => img.idx !== idx,
        );
      }

      if (genreText) {
        chosenBlock.genres = chosenBlock.genres!.filter(
          (genre) => genre !== genreText,
        );
      }

      if (deleteBlock) {
        state[i].data = state[i].data.filter(
          (block) => block.blockID !== blockID,
        );
      }
    },
  },
});

export const selectDatabaseData = (state: RootState) => state.databaseData;
export const {
  populateDatabaseData,
  updateDatabaseData,
  addToDatabaseData,
  removeFromDatabaseData,
  clearDatabaseData,
} = databaseDataSlice.actions;
export default databaseDataSlice.reducer;
