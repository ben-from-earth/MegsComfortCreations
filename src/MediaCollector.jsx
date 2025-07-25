import BackgroundIMG from "./assets/FlowerBackground.png";
import MediaCollectorTitle from "./assets/MegsMediaCollector.png";
import { useEffect, useReducer } from "react";
import "./MediaCollector.css";
import MediaInputs from "./MediaInputs";
import ButtonGroup from "./ButtonGroup";
import DataContext from "./DataContext";

const API_KEY = "AIzaSyCtoXKuRUP5p0Xrk21635t67OA6MxFLay4";
const CX = "e1e30c1aaa513492b";

function reducer(state, action) {
  switch (action.type) {
    case "set-checks":
      return {
        ...state,
        mediaTypes: state.mediaTypes.map((mediaType, i) =>
          action.idx === i ? { ...mediaType, show: !mediaType.show } : mediaType
        ),
      };
    case "set-collect-text": {
      let searchArr = action.text.split("/").map((i) => i.trim());
      searchArr = searchArr.map((t) =>
        t
          ? t +
            ` ${action.mediaType.slice(
              0,
              action.mediaType.length - 1
            )} Cover Image`
          : ""
      );
      return {
        ...state,
        mediaTypes: state.mediaTypes.map((mediaType) =>
          action.mediaType === mediaType.type
            ? { ...mediaType, titles: searchArr }
            : mediaType
        ),
      };
    }
    case "Collect": {
      return { ...state, shouldFetch: true, isLoading: true };
    }
    case "Finished Fetch": {
      console.log(action.payload);
    }
  }
}

const MediaCollector = () => {
  const medias = ["Books", "Movies", "Video Games", "Albums"];
  const [Data, dispatch] = useReducer(reducer, {
    mediaTypes: medias.map((m) => ({ type: m, show: false, titles: [] })),
    shouldFetch: false,
    isLoading: false,
  });

  useEffect(() => {
    if (!Data.shouldFetch) return;

    const controller = new AbortController();
    async function CollectMediaCovers(title) {
      const params = new URLSearchParams({
        q: title,
        cx: CX,
        key: API_KEY,
        searchType: "image",
        num: 3,
      });
      try {
        const res = await fetch(
          `https://www.googleapis.com/customsearch/v1?${params.toString()}`,
          { signal: controller.signal }
        );
        const data = await res.json();
        dispatch({ type: "Finished Fetch", payload: data });
      } catch (e) {
        if (e.name !== "Abort Error") {
          console.log(e);
        }
      }
    }
    Data.mediaTypes.map(({ titles }) =>
      titles.map((t) => CollectMediaCovers(t))
    );

    return () => controller.abort();
  }, [Data.shouldFetch]);

  return (
    <DataContext.Provider value={{ dispatch }}>
      <div
        className="InfoForm"
        style={{
          backgroundImage: `url(${BackgroundIMG})`,
        }}
      >
        <img src={`${MediaCollectorTitle}`} />
        <MediaInputs info={Data} />
        <ButtonGroup />
      </div>
    </DataContext.Provider>
  );
};

export default MediaCollector;
