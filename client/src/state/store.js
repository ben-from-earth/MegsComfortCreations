import { configureStore } from '@reduxjs/toolkit';
import collectorReducer from '@/state/collectorSlice';
import databaseDataReducer from '@/state/databaseDataSlice';
import pngCollectionReducer from '@/state/pngCollectionSlice';

export const store = configureStore({
  reducer: {
    collector: collectorReducer,
    databaseData: databaseDataReducer,
    pngCollection: pngCollectionReducer,
  },
  devTools: true,
});
