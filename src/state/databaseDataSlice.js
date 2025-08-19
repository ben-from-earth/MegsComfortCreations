import { createAsyncThunk, createSlice } from "@reduxjs/toolkit";
import { medias } from "./collectorSlice";

const initialState = medias.map(({ type, label }) => ({
  type,
  label,
  data: [],
}));

export const sendToDatabase = createAsyncThunk(
  "/databaseData/sendtodb",
  async ({ databaseData }) => {
    for (const media of databaseData) {
      const sendData = media.data;
      if (media.type === "book") {
        const bookPromises = sendData.map(async (book) => {
          //bookData looks like {title, author, pub_year, page_count, blockID, images:[{src, idx}]}
          //need {title, author, page_count, pub_year, image_urls}
          const bookData = {
            title: book.title,
            author: book.author,
            page_count: book.page_count,
            pub_year: book.pub_year,
            image_urls: book.images.map((item) => item.src),
          };

          const res = await fetch("http://localhost:3001/savetodb/book", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(bookData),
          });

          if (!res.ok) {
            throw new Error(`Server Error ${res.status}`);
          }
          return res.json();
        });

        try {
          // Wait for all to finish
          const results = await Promise.allSettled(bookPromises);
          results.forEach((r) => {
            if (r.status === "fulfilled") {
              console.log("Server response:", r.value);
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

          const res = await fetch(
            `http://localhost:3001/savetodb/${media.type}`,
            {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify(mediaData),
            }
          );

          if (!res.ok) {
            throw new Error(`Server Error ${res.status}`);
          }
          return res.json();
        });

        try {
          // Wait for all to finish
          const results = await Promise.allSettled(mediaPromises);
          results.forEach((r) => {
            if (r.status === "fulfilled") {
              console.log("Server response:", r.value);
            } else {
              console.error("Save failed:", r.reason);
            }
          });
        } catch (err) {
          console.error(
            `Unexpected error saving ${media.label.toLowerCase()}s:`,
            err
          );
        }
      }
    }
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
    addImageToDatabaseData: (state, action) => {
      const { blockID, type, idx, src } = action.payload;
      const i = state.findIndex((m) => m.type === type);
      const j = state[i].data.findIndex((block) => block.blockID === blockID);

      state[i].data[j].images.push({ src, idx });
    },
    //remove image from state storage if user deselects it
    removeImageFromDatabaseData: (state, action) => {
      const { blockID, type, idx } = action.payload;
      const i = state.findIndex((m) => m.type === type);
      const j = state[i].data.findIndex((block) => block.blockID === blockID);
      state[i].data[j].images = state[i].data[j].images.filter(
        (img) => img.idx !== idx
      );
    },
  },
});

export const selectDatabaseData = (state) => state.databaseData;
export const {
  populateDatabaseData,
  updateDatabaseData,
  addImageToDatabaseData,
  removeImageFromDatabaseData,
  clearDatabaseData,
} = databaseDataSlice.actions;
export default databaseDataSlice.reducer;
