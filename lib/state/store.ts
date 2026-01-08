import { configureStore } from '@reduxjs/toolkit';
import pngCollectionReducer from 'lib/state/slices/pngCollectionSlice';
import { useDispatch } from 'react-redux';

export const store = configureStore({
  reducer: {
    pngCollection: pngCollectionReducer,
  },
  devTools: true,
});

export type RootState = ReturnType<typeof store.getState>;
export type AppDispatch = typeof store.dispatch;
export const useAppDispatch = useDispatch.withTypes<AppDispatch>();
