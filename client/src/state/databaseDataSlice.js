import { createAsyncThunk, createSlice } from "@reduxjs/toolkit";
import { medias } from "./collectorSlice";

const initialState = medias.map(({ type, label }) => ({
  type,
  label,
  data: [],
}));

const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

export const sendToDatabase = createAsyncThunk(
  "/databaseData/sendtodb",
  async ({ databaseData }) => {
    const serverResponses = [];
    for (const media of databaseData) {
      const sendData = media.data;
      if (media.type === "book") {
        const bookPromises = sendData.map(async (book) => {
          //bookData looks like {title, author, pub_year, page_count, blockID, images:[{src, idx}], genres: [], spine_color}
          //need {title, author, page_count, pub_year, image_urls, spine_color}
          const bookData = {
            title: book.title,
            author: book.author,
            page_count: book.page_count,
            pub_year: book.pub_year,
            image_urls: book.images.map((item) => item.src),
            spine_color: book.spine_color,
          };

          const bookSaveRes = await fetch(
            `${serverDomain}/database/save/book`,
            {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify(bookData),
            }
          );

          if (!bookSaveRes.ok) {
            throw new Error(`Server Error ${bookSaveRes.status}`);
          }
          const bookCreationResponse = await bookSaveRes.json();
          const bookDatabaseID = bookCreationResponse.saved_book.id;

          const genreLinkRes = await fetch(`${serverDomain}/genres/addLink`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              bookID: bookDatabaseID,
              genres: book.genres,
            }),
          });
          if (!genreLinkRes.ok) {
            throw new Error(`Server Error ${genreLinkRes.status}`);
          }

          const genreLinkResponse = await genreLinkRes.json();

          return { ...bookCreationResponse, ...genreLinkResponse };
        });

        try {
          // Wait for all to finish
          const results = await Promise.allSettled(bookPromises);
          results.forEach((r) => {
            if (r.status === "fulfilled") {
              serverResponses.push(r.value);
            } else {
              console.error("Save failed:", r.reason);
            }
          });
        } catch (err) {
          console.error("Unexpected error saving books:", err);
        }
      } else {
        const mediaPromises = sendData.map(async (item) => {
          //mediaData looks like {title, blockID, images:[{src, idx}]}
          //need {title, image_urls}
          const mediaData = {
            title: item.title,
            image_urls: item.images.map((item) => item.src),
          };

          try {
            const res = await fetch(
              `${serverDomain}/database/save/${media.type}`,
              {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(mediaData),
              }
            );

            if (!res.ok) {
              //errors are correct captured from the server here
              return res.json();
            }
            return res.json();
          } catch (error) {
            return {
              error: "Server Error",
              message: "Something went wrong connecting to the server",
            };
          }
        });

        // Wait for all to finish
        const results = await Promise.allSettled(mediaPromises);

        serverResponses.push(...results.map((result) => result.value));
      }
    }
    return serverResponses;
  }
);

export const databaseDataSlice = createSlice({
  name: "databaseData",
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
          type === "book" &&
          item.title === data.title &&
          item.author === data.author
        ) {
          exists = true;
        } else if (type !== "book" && item.title === data.title) {
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
        name === "pub_year" || name === "page_count"
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
          (img) => img.idx !== idx
        );
      }

      if (genreText) {
        chosenBlock.genres = chosenBlock.genres.filter(
          (genre) => genre !== genreText
        );
      }

      if (deleteBlock) {
        state[i].data = state[i].data.filter(
          (block) => block.blockID !== blockID
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
