import { configureStore } from "@reduxjs/toolkit";
import collectorReducer from "./collectorSlice";
import databaseDataReducer from "./databaseDataSlice";
import pngCollectionReducer from "./pngCollectionSlice";

export const store = configureStore({
  reducer: {
    collector: collectorReducer,
    databaseData: databaseDataReducer,
    pngCollection: pngCollectionReducer,
  },
  devTools: true,
});
