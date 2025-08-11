import { configureStore } from "@reduxjs/toolkit";
import collectorReducer from "./collectorSlice";
import databaseDataReducer from "./databaseDataSlice";

export const store = configureStore({
  reducer: {
    collector: collectorReducer,
    databaseData: databaseDataReducer,
  },
  devTools: true,
});
