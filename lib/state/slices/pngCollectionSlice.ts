// react, redux imports
import { createSlice, PayloadAction } from '@reduxjs/toolkit';
import { RootState } from 'lib/state/store';

// interfaces and types
export interface ImageData {
  url: string;
  type: 'book' | 'movie' | 'videoGame' | 'album';
  spineColor: string;
}

const initialState = {
  pngCollectionList: [] as ImageData[],
};

export const pngCollectionSlice = createSlice({
  name: 'pngCollection',
  initialState,
  reducers: {
    addToPNGCollectionList: (state, action: PayloadAction<ImageData>) => {
      const ImageData: ImageData = action.payload;
      if (!state.pngCollectionList.some((item) => ImageData.url === item.url)) {
        state.pngCollectionList.push(ImageData);
      }
    },
    removeFromPNGCollectionList: (
      state,
      action: PayloadAction<{ url: string }>,
    ) => {
      state.pngCollectionList = state.pngCollectionList.filter(
        (image) => image.url !== action.payload.url,
      );
    },
    clearPNGCollectionList: (state) => {
      state.pngCollectionList = [];
    },
  },
});

export const selectPNGList = (state: RootState) =>
  state.pngCollection.pngCollectionList;
export const {
  addToPNGCollectionList,
  removeFromPNGCollectionList,
  clearPNGCollectionList,
} = pngCollectionSlice.actions;

export default pngCollectionSlice.reducer;
