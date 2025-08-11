import { createSlice } from "@reduxjs/toolkit";
import { medias } from "./collectorSlice";

const initialState = medias.map(({ id, label }) => ({
  id,
  label,
  data: [],
}));

export const databaseDataSlice = createSlice({
  name: "databaseData",
  initialState,
  reducers: {
    populateDatabaseData: (state, action) => {
      const { type, data } = action.payload;
      const i = state.findIndex((m) => m.id === type); //this is an array of objects described in the initial state
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
      const { id, type, name, newText } = action.payload;
      const i = state.findIndex((m) => m.id === type);
      const j = state[i].data.findIndex((block) => block.blockID === id);
      state[i].data[j][name] =
        name === "first_publish_year" || name === "number_of_pages"
          ? Number(newText)
          : newText;
    },
    addImageToDatabaseData: (state, action) => {
      const { id, type, idx, text } = action.payload;
      const i = state.findIndex((m) => m.id === type);
      const j = state[i].data.findIndex((block) => block.blockID === id);

      state[i].data[j].images.push({ text, idx });
    },
    removeImageFromDatabaseData: (state, action) => {
      const { id, type, idx } = action.payload;
      const i = state.findIndex((m) => m.id === type);
      const j = state[i].data.findIndex((block) => block.blockID === id);
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
} = databaseDataSlice.actions;
export default databaseDataSlice.reducer;
