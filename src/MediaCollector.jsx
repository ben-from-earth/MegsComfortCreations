import BackgroundIMG from "./assets/FlowerBackground.png";
import MediaCollectorTitle from "./assets/MegsMediaCollector.png";
import { useEffect, useReducer, useRef, useState } from "react";
import "./MediaCollector.css";
import MediaInputs from "./MediaInputs";
import MediaCheckboxes from "./MediaCheckboxes";
import ButtonGroup from "./ButtonGroup";
import DataContext from "./DataContext";
import MediaDataContext from "./MediaDataContext";
import { v4 as uuid } from "uuid";
import TitleBlockContainer from "./TitleBlockContainer";
import QueryCounter from "./QueryCounter";

const API_KEY = import.meta.env.VITE_GOOGLE_API_KEY;
const CX = import.meta.env.VITE_CX;

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
      let searchArr = action.text.split(",").map((i) => i.trim());
      searchArr = searchArr.map((t) => {
        const titleInfo = t.split("/").map((i) => i.trim());
        const title = titleInfo[0];
        const author = titleInfo[1];

        return {
          title,
          author,
        };
      });
      return {
        ...state,
        mediaTypes: state.mediaTypes.map((mediaType) =>
          action.mediaType === mediaType.type
            ? { ...mediaType, toCollect: searchArr }
            : mediaType
        ),
      };
    }
    case "Collect": {
      return { ...state, shouldFetch: true, isLoading: true };
    }
    case "Finished Fetch":
      return {
        ...state,
        shouldFetch: false,
        isLoading: false,
        mediaTypes: state.mediaTypes.map((mt) => ({
          ...mt,
          toCollect: [],
        })),
      };
    case "set-toCollect-data": {
      return state;
    }
    case "send-to-database": {
      console.log(action.items);
      return state;
    }
  }
}

const MediaCollector = () => {
  const medias = ["Book", "Movie", "Video Game", "Album"];
  const [Data, dispatch] = useReducer(reducer, {
    mediaTypes: medias.map((m) => ({ type: m, show: false, toCollect: [] })),
    shouldFetch: false,
    isLoading: false,
  });
  const [CollectedCoversBlocks, setCollectedCoversBlocks] = useState([]);

  const updateQueryCount = () => {
    const today = new Date().toISOString().split("T")[0];
    const storedDate = localStorage.getItem("lastQueryDate");

    if (storedDate !== today) {
      localStorage.setItem("queryCount", "0");
      localStorage.setItem("lastQueryDate", today);
    }

    let qCount = Number(localStorage.getItem("queryCount"));
    qCount++;
    localStorage.setItem("queryCount", `${qCount}`);
  };

  const mediaTypesRef = useRef(Data.mediaTypes);

  useEffect(() => {
    mediaTypesRef.current = Data.mediaTypes;
  }, [Data.mediaTypes]);

  useEffect(() => {
    if (!Data.shouldFetch) return;

    let isCancelled = false;
    setCollectedCoversBlocks([]);

    async function grabOpenLibraryData({ title, author }) {
      try {
        const URL = `https://openlibrary.org/search.json?title=${title.replaceAll(
          " ",
          "+"
        )}&author=${author.replaceAll(" ", "+")}&limit=1`;
        const res = await fetch(URL);
        const data = await res.json();
        const {
          docs: [{ first_publish_year, cover_edition_key }],
        } = data;

        const editionURL = `https://openlibrary.org/books/${cover_edition_key}.json`;
        const editionRes = await fetch(editionURL);
        const editionData = await editionRes.json();

        const { number_of_pages } = editionData;
        return { title, author, first_publish_year, number_of_pages };
      } catch {
        console.log(
          `Error in OpenLibrary API Call, data not collected for ${title}`
        );
        return { title, author };
      }
    }

    async function CollectMediaCovers(type, { title, author }) {
      const controller = new AbortController();

      const imgArr = [];
      try {
        const params = new URLSearchParams({
          q: `${title} ${type} Cover Image`,
          cx: CX,
          key: API_KEY,
          searchType: "image",
          num: 3,
        });

        const res = await fetch(
          `https://www.googleapis.com/customsearch/v1?${params.toString()}`,
          { signal: controller.signal }
        );

        if (!res.ok) {
          throw new Error(
            `Google Search API failed: ${res.status} ${res.statusText}`
          );
        }
        const data = await res.json();
        updateQueryCount();
        data.items.map((i) => imgArr.push(i.link));
      } catch (e) {
        if (e.name !== "AbortError") {
          console.log(e.message);
        } else {
          console.log("Aborted");
        }
      }
      let blockInfo;
      if (type === "Book") {
        blockInfo = await grabOpenLibraryData({ title, author }); // { title, author, first_publish_year, number_of_pages }
      } else {
        blockInfo = { title };
      }
      dispatch({ type: "set-toCollect-data", payload: [type, blockInfo] });

      setCollectedCoversBlocks((blocks) => [
        ...blocks,
        { type, images: imgArr, blockInfo, id: uuid() },
      ]);
    }

    const work = mediaTypesRef.current
      .filter((mt) => mt.toCollect.length > 0)
      .flatMap(({ type, toCollect }) =>
        toCollect.map((t) => ({ type, payload: t }))
      );

    const promises = work.map(({ type, payload }) =>
      CollectMediaCovers(type, payload)
    );

    Promise.all(promises)
      .then(() => {
        if (isCancelled) return;
        dispatch({ type: "Finished Fetch" });
      })
      .catch(console.error);

    return () => {
      isCancelled = true;
    };
  }, [Data.shouldFetch]);

  return (
    <>
      <DataContext.Provider value={{ dispatch }}>
        <div
          className="InfoForm"
          style={{
            backgroundImage: `url(${BackgroundIMG})`,
          }}
        >
          <QueryCounter />
          <img src={`${MediaCollectorTitle}`} />
          <MediaCheckboxes mediaTypes={Data.mediaTypes} />
          <MediaDataContext.Provider value={{ CollectedCoversBlocks }}>
            <ButtonGroup />
          </MediaDataContext.Provider>
          <MediaInputs mediaTypes={Data.mediaTypes} />
        </div>
      </DataContext.Provider>
      {CollectedCoversBlocks.length > 0 && (
        <TitleBlockContainer blocks={CollectedCoversBlocks} />
      )}
      {Data.isLoading && <p>Loading...</p>}
    </>
  );
};

export default MediaCollector;
