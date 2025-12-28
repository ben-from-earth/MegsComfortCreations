// react, redux imports
import { createAsyncThunk, createSlice, PayloadAction } from '@reduxjs/toolkit';
import { RootState } from 'lib/state/store';

// necessary imports from collector state
import { medias } from 'lib/state/slices/collectorSlice';

// helpers
import { titleRearrange } from 'lib/helpers/titleRearrange';

// interfaces and types
import { DatabasePayload } from '@//mediacollector/CollectedCoversBlock';
import {
  AlbumInsert,
  BookInsert,
  DatabaseSaveServerResponse,
  MediaType,
  MovieInsert,
  PreSavedMediaItem,
  SuccessfulGenreLinkUnlinkResponse,
  SuccessfulMediaSaveEditResponse,
  VideoGameInsert,
} from 'lib/interfaces/globalInterfaces';
import { DatabaseSaveEditErrorResponse } from '@//api/api-Errors';
import { trpcClient } from 'lib/trpc/vanillaClient';

type WithImages<T> = T & {
  images: { src: string; idx: number }[];
};

export type DatabaseSliceDataBook = WithImages<BookInsert>;
export type DatabaseSliceDataMovie = WithImages<MovieInsert>;
export type DatabaseSliceDataVideoGame = WithImages<VideoGameInsert>;
export type DatabaseSliceDataAlbum = WithImages<AlbumInsert>;

export type DatabaseDataPerType =
  | {
      type: 'book';
      label: 'Book';
      data: DatabaseSliceDataBook[];
    }
  | {
      type: 'movie';
      label: 'Movie';
      data: DatabaseSliceDataMovie[];
    }
  | {
      type: 'videoGame';
      label: 'Video Game';
      data: DatabaseSliceDataVideoGame[];
    }
  | {
      type: 'album';
      label: 'Album';
      data: DatabaseSliceDataAlbum[];
    };

//set up initial state
const initialState: DatabaseDataPerType[] = medias.map((m) => ({
  ...m,
  data: [] as never[],
})) as DatabaseDataPerType[];

export const sendToDatabase = createAsyncThunk(
  '/databaseData/sendtodatabase',
  async (
    databaseData: DatabaseDataPerType[],
  ): Promise<DatabaseSaveServerResponse> => {
    const serverResponses: DatabaseSaveServerResponse = [];
    for (const media of databaseData) {
      if (media.type === 'book') {
        const sendData = media.data;
        const bookPromises = sendData.map(async (book) => {
          const title = titleRearrange(book.title);
          const bookData = {
            title,
            author: book.author,
            pageCount: book.pageCount,
            pubYear: book.pubYear,
            imageUrls: book.images.map((item) => item.src),
            spineColor: book.spineColor,
          };

          try {
            const bookCreationResponse = (await trpcClient.database.save.mutate(
              {
                type: 'book',
                mediaData: bookData,
              },
            )) as
              | SuccessfulMediaSaveEditResponse
              | DatabaseSaveEditErrorResponse;

            if ('error' in bookCreationResponse) {
              return bookCreationResponse;
            } else {
              const bookDatabaseID = bookCreationResponse.actionAttemptItem.id;
              const genreLinkResponse = (await trpcClient.genres.link.mutate({
                bookID: bookDatabaseID,
                genres: book.genres || [],
              })) as { genreResponses: SuccessfulGenreLinkUnlinkResponse[] };
              return { ...bookCreationResponse, ...genreLinkResponse };
            }
          } catch {
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
        const sendData = media.data;
        const otherMediaPromises = sendData.map(async (otherMedia) => {
          const title = titleRearrange(otherMedia.title);
          const otherMediaData: PreSavedMediaItem = {
            title,
            imageUrls: otherMedia.images.map((img) => img.src),
            spineColor: otherMedia.spineColor,
          };

          try {
            const otherMediaCreationResponse =
              await trpcClient.database.save.mutate({
                type: media.type,
                mediaData: otherMediaData,
              });
            return otherMediaCreationResponse;
          } catch {
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
        });
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
      action: PayloadAction<DatabasePayload>,
    ): void => {
      const payload = action.payload;
      const i = state.findIndex((m) => m.type === payload.type);
      if (i === -1) return;

      const mediaState = state[i];

      if (payload.type === 'book' && mediaState.type === 'book') {
        const bookDataArray = mediaState.data;
        const bookData = payload.data;

        let exists = false;
        for (const item of bookDataArray) {
          // item: DatabaseSliceDataBook
          if (
            item.title === bookData.title &&
            item.author === bookData.author
          ) {
            exists = true;
            break;
          }
        }

        if (!exists) {
          mediaState.data.push({ ...bookData, images: [] });
        }
      } else if (payload.type !== 'book' && mediaState.type !== 'book') {
        const otherData = payload.data;

        let exists = false;
        for (const item of mediaState.data) {
          if (item.title === otherData.title) {
            exists = true;
            break;
          }
        }

        if (!exists) {
          mediaState.data.push({ ...otherData, images: [] });
        }
      }
    },
    clearDatabaseData: () => initialState,

    // if the user updates the data in the collectedCoverBlock text area,
    //reflect the changes in the state as well
    updateDatabaseData: (
      state,
      action: PayloadAction<{
        blockID: string;
        type: 'book' | 'movie' | 'videoGame' | 'album';
        name: 'title' | 'author' | 'pubYear' | 'pageCount';
        newText: string;
      }>,
    ) => {
      const payload = action.payload;

      // find which media bucket we're editing
      const mediaIndex = state.findIndex((m) => m.type === payload.type);
      if (mediaIndex === -1) return;

      const mediaState = state[mediaIndex];

      // find which block within that bucket
      const blockIndex = mediaState.data.findIndex(
        (block) => block.blockID === payload.blockID,
      );
      if (blockIndex === -1) return;

      // ---- BOOK BRANCH --------------------------------------------------------
      if (mediaState.type === 'book') {
        // TS + Immer sometimes still won't narrow `mediaState.data[blockIndex]`,
        // so we help it with a one-time cast that is logically safe because of the guard:
        const block = mediaState.data[blockIndex] as DatabaseSliceDataBook;

        if (payload.name === 'title') {
          block.title = payload.newText;
        } else if (payload.name === 'author') {
          block.author = payload.newText;
        } else if (payload.name === 'pubYear') {
          block.pubYear = Number(payload.newText);
        } else if (payload.name === 'pageCount') {
          block.pageCount = Number(payload.newText);
        }

        return;
      }

      // ---- NON-BOOK BRANCH ----------------------------------------------------
      // Here mediaState.type is 'movie' | 'videoGame' | 'album'
      // These types only have `title`, so we only handle that safely.
      const block = mediaState.data[blockIndex];

      if (payload.name === 'title') {
        block.title = payload.newText;
      }
      // ignore author / pubYear / pageCount for non-book types
    },

    //push image to state storage if user selects it
    addToDatabaseData: (
      state,
      action: PayloadAction<{
        blockID: string;
        type: MediaType;
        idx?: number;
        src?: string;
        spineColor?: string;
        genreText?: string;
      }>,
    ) => {
      const { blockID, type, idx, src, spineColor, genreText } = action.payload;
      const i = state.findIndex((m) => m.type === type);
      const j = state[i].data.findIndex((block) => block.blockID === blockID);
      const chosenBlock = state[i].data[j];
      if (src) chosenBlock.images.push({ src, idx: idx! });
      if (spineColor) chosenBlock.spineColor = spineColor;
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
