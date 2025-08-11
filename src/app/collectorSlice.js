import { createSlice } from "@reduxjs/toolkit";

export const medias = [
  { id: "book", label: "Book" },
  { id: "movie", label: "Movie" },
  { id: "video_game", label: "Video Game" },
  { id: "album", label: "Album" },
];
const initialState = {
  mediaTypes: medias.map(({ id, label }) => ({
    type: id,
    label,
    show: false,
    toCollect: [],
  })),
  shouldFetch: false,
  isLoading: false,
};

export const collectorSlice = createSlice({
  name: "collector",
  initialState,
  reducers: {
    setChecks: (state, action) => {
      const idx = action.payload;
      state.mediaTypes[idx].show = !state.mediaTypes[idx].show;
    },
    setCollectText: (state, action) => {
      for (let media of action.payload.searchData) {
        let searchArr = [];
        if (media.type === "book" && media.text) {
          searchArr = media.text.split(",").map((i) => i.trim());
          searchArr = searchArr.map((t) => {
            const titleInfo = t.split("/").map((i) => i.trim());
            const title = titleInfo[0];
            const author = titleInfo[1];

            return {
              title,
              author,
            };
          });
        } else if (media.text) {
          searchArr = media.text.split(",").map((i) => i.trim());
        }
        const i = state.mediaTypes.findIndex((m) => m.type === media.type);
        if (i !== -1) state.mediaTypes[i].toCollect = searchArr;
      }
    },
    collectMedia: (state) => {
      state.shouldFetch = true;
      state.isLoading = true;
    },
    finishedFetch: (state) => {
      state.isLoading = false;
      state.shouldFetch = false;
      state.mediaTypes = state.mediaTypes.map((mt) => ({
        ...mt,
        toCollect: [],
      }));
    },
  },
});

export const mediaData = (state) => state.collector.mediaTypes;
export const getFetchStatus = (state) => state.collector.shouldFetch;
export const getLoadingStatus = (state) => state.collector.isLoading;
export const { setChecks, setCollectText, collectMedia, finishedFetch } =
  collectorSlice.actions;
export default collectorSlice.reducer;
