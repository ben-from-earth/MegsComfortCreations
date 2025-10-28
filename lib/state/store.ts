import { configureStore } from "@reduxjs/toolkit";
import collectorReducer from "@/lib/state/slices/collectorSlice";
import databaseDataReducer from "@/lib/state/slices/databaseDataSlice";
import pngCollectionReducer from "@/lib/state/slices/pngCollectionSlice";
import { useDispatch } from "react-redux";

export const store = configureStore({
  reducer: {
    collector: collectorReducer,
    databaseData: databaseDataReducer,
    pngCollection: pngCollectionReducer,
  },
  devTools: true,
});

export type RootState = ReturnType<typeof store.getState>;
export type AppDispatch = typeof store.dispatch;
export const useAppDispatch = useDispatch.withTypes<AppDispatch>();
