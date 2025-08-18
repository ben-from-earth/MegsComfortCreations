import { createSlice } from "@reduxjs/toolkit";
import { medias } from "./collectorSlice";

const initialState = medias.map(({ type, label }) => ({
  type,
  label,
  data: [],
}));

export const databaseDataSlice = createSlice({
  name: "databaseData",
  initialState,
  reducers: {
    populateDatabaseData: (state, action) => {
      const { type, data } = action.payload;
      const i = state.findIndex((m) => m.type === type); //this is an array of objects described in the initial state
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
    updateDatabaseData: (state, action) => {
      const { blockID, type, name, newText } = action.payload;
      const i = state.findIndex((m) => m.type === type);
      const j = state[i].data.findIndex((block) => block.blockID === blockID);
      state[i].data[j][name] =
        name === "first_publish_year" || name === "number_of_pages"
          ? Number(newText)
          : newText;
    },
    addImageToDatabaseData: (state, action) => {
      const { blockID, type, idx, src } = action.payload;
      const i = state.findIndex((m) => m.type === type);
      const j = state[i].data.findIndex((block) => block.blockID === blockID);

      state[i].data[j].images.push({ src, idx });
    },
    removeImageFromDatabaseData: (state, action) => {
      const { blockID, type, idx } = action.payload;
      const i = state.findIndex((m) => m.type === type);
      const j = state[i].data.findIndex((block) => block.blockID === blockID);
      state[i].data[j].images = state[i].data[j].images.filter(
        (img) => img.idx !== idx
      );
    },
    sendToDatabase: (state, action) => {
      const databaseData = action.payload.databaseData;
      databaseData.forEach((media) => {
        const sendData = media.data;
        if (media.type === "book") {
          sendData.forEach((book) => {
            console.log(book);
          });
        } else {
          sendData.forEach((media) => {
            const refactor = { title: media.title, images: media.images };
            console.log(refactor);
          });
        }
      });
    },
  },
});

export const selectDatabaseData = (state) => state.databaseData;
export const {
  populateDatabaseData,
  updateDatabaseData,
  addImageToDatabaseData,
  removeImageFromDatabaseData,
  sendToDatabase,
} = databaseDataSlice.actions;
export default databaseDataSlice.reducer;
