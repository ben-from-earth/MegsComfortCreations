//library imports
import axios from 'axios';

//redux imports
import { createAsyncThunk, createSlice } from '@reduxjs/toolkit';

//imports from collector state slice
import { medias } from './collectorSlice';

//import helper function for title rearranging
import { titleRearrange } from '../pages/MediaCollector/helpers/mediaCollectorHelpers';

const initialState = medias.map(({ type, label }) => ({
  type,
  label,
  data: [],
}));

const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

export const sendToDatabase = createAsyncThunk(
  '/databaseData/sendtodb',
  async ({ databaseData }) => {
    const serverResponses = [];
    for (const media of databaseData) {
      const sendData = media.data;
      if (media.type === 'book') {
        const bookPromises = sendData.map(async (book) => {
          //bookData looks like {title, author, pub_year, page_count, blockID, images:[{src, idx}], genres: [], spine_color}
          //need {title, author, page_count, pub_year, image_urls, spine_color}

          //Do a bit of syntax rearranging to books starting with A, An, or The
          const title = titleRearrange(book.title);

          const bookData = {
            title,
            author: book.author,
            page_count: book.page_count,
            pub_year: book.pub_year,
            image_urls: book.images.map((item) => item.src),
            spine_color: book.spine_color,
          };

          try {
            const bookSaveRes = await axios.post(
              `${serverDomain}/database/save/book`,
              bookData,
              { validateStatus: (status) => status < 500 },
            );

            const bookCreationResponse = bookSaveRes.data;
            if (bookCreationResponse.saved === false) {
              return { ...bookCreationResponse };
            } else {
              const bookDatabaseID = bookCreationResponse.saveAttemptItem.id;

              const genreLinkRes = await axios.post(
                `${serverDomain}/genres/addLink`,
                {
                  bookID: bookDatabaseID,
                  genres: book.genres,
                },
                { validateStatus: (status) => status < 500 },
              );
              const genreLinkResponse = genreLinkRes.data;

              return { ...bookCreationResponse, ...genreLinkResponse };
            }
          } catch (error) {
            return {
              error: 'Server Error',
              message: 'Something went wrong connecting to the server',
            };
          }
        });
        const results = await Promise.allSettled(bookPromises);
        serverResponses.push(...results.map((result) => result.value));
      } else {
        const mediaPromises = sendData.map(async (item) => {
          //mediaData looks like {title, blockID, images:[{src, idx}]}
          //need {title, image_urls}

          const mediaData = {
            title: item.title,
            image_urls: item.images.map((item) => item.src),
            spine_color: item.spine_color,
          };

          try {
            const res = await axios.post(
              `${serverDomain}/database/save/${media.type}`,
              mediaData,
              { validateStatus: (status) => status < 500 },
            );

            return res.data;
          } catch (error) {
            return {
              error: 'Server Error',
              message: 'Something went wrong connecting to the server',
            };
          }
        });

        // Wait for all to finish
        const results = await Promise.allSettled(mediaPromises);

        serverResponses.push(...results.map((result) => result.value));
      }
    }
    return serverResponses;
  },
);

export const databaseDataSlice = createSlice({
  name: 'databaseData',
  initialState,
  reducers: {
    //populate the database data with the data used to create the collectCoversBlocks
    populateDatabaseData: (state, action) => {
      const { type, data } = action.payload;
      const i = state.findIndex((m) => m.type === type);

      //check if the data already exists before pushing
      let exists = false;
      for (let item of state[i].data) {
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
      if (exists === false) {
        state[i].data.push({ ...data, images: [] });
      }
    },
    clearDatabaseData: () => initialState,

    // if the user updates the data in the collectedCoverBlock text area,
    //reflect the changes in the state as well
    updateDatabaseData: (state, action) => {
      const { blockID, type, name, newText } = action.payload;
      const i = state.findIndex((m) => m.type === type);
      const j = state[i].data.findIndex((block) => block.blockID === blockID);
      state[i].data[j][name] =
        name === 'pub_year' || name === 'page_count'
          ? Number(newText)
          : newText;
    },
    //push image to state storage if user selects it
    addToDatabaseData: (state, action) => {
      const { blockID, type, idx, src, spine_color, genreText } =
        action.payload;
      const i = state.findIndex((m) => m.type === type);
      const j = state[i].data.findIndex((block) => block.blockID === blockID);
      const chosenBlock = state[i].data[j];
      if (src) chosenBlock.images.push({ src, idx });
      if (spine_color) chosenBlock.spine_color = spine_color;
      if (genreText) chosenBlock.genres.push(genreText);
    },
    //remove image from state storage if user deselects it
    removeFromDatabaseData: (state, action) => {
      const { blockID, type, idx, genreText, deleteBlock } = action.payload;
      const i = state.findIndex((m) => m.type === type);
      const j = state[i].data.findIndex((block) => block.blockID === blockID);
      const chosenBlock = state[i].data[j];
      if (isFinite(idx)) {
        chosenBlock.images = chosenBlock.images.filter(
          (img) => img.idx !== idx,
        );
      }

      if (genreText) {
        chosenBlock.genres = chosenBlock.genres.filter(
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

export const selectDatabaseData = (state) => state.databaseData;
export const {
  populateDatabaseData,
  updateDatabaseData,
  addToDatabaseData,
  removeFromDatabaseData,
  clearDatabaseData,
} = databaseDataSlice.actions;
export default databaseDataSlice.reducer;
