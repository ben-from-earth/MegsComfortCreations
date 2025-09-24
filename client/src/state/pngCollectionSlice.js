//redux
import { createSlice } from '@reduxjs/toolkit';

const initialState = {
  pngCollectionList: [],
};

export const pngCollectionSlice = createSlice({
  name: 'pngCollection',
  initialState,
  reducers: {
    addToPNGCollectionList: (state, action) => {
      //action.payload = {url, type, spine_color}
      if (
        !state.pngCollectionList.some((item) => action.payload.url === item.url)
      ) {
        state.pngCollectionList.push(action.payload);
      }
    },
    removeFromPNGCollectionList: (state, action) => {
      //action.payload = {url}
      state.pngCollectionList = state.pngCollectionList.filter(
        (image) => image.url !== action.payload.url,
      );
    },
    clearPNGCollectionList: (state) => {
      state.pngCollectionList = [];
    },
  },
});

export const selectPNGList = (state) => state.pngCollection.pngCollectionList;
export const {
  addToPNGCollectionList,
  removeFromPNGCollectionList,
  clearPNGCollectionList,
} = pngCollectionSlice.actions;

export default pngCollectionSlice.reducer;
