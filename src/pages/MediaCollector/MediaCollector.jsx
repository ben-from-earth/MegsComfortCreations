import BackgroundIMG from "../../assets/FlowerBackground.png";
import MediaCollectorTitle from "../../assets/MegsMediaCollector.png";
import { useEffect, useRef, useState } from "react";
import "./MediaCollector.css";
import MediaInputs from "./MediaInputs";
import MediaCheckboxes from "./MediaCheckboxes";
import ButtonGroup from "./ButtonGroup";
import TitleBlockContainer from "./TitleBlockContainer";
import QueryCounter from "./QueryCounter";
import { useDispatch, useSelector } from "react-redux";
import {
  finishedFetch,
  getFetchStatus,
  getLoadingStatus,
  mediaData,
} from "../../app/collectorSlice";
import { nanoid } from "@reduxjs/toolkit";

const API_KEY = import.meta.env.VITE_GOOGLE_API_KEY;
const CX = import.meta.env.VITE_CX;

const MediaCollector = () => {
  const dispatch = useDispatch();
  const [CollectedCoversBlocks, setCollectedCoversBlocks] = useState([]);

  const Data = useSelector(mediaData);
  const shouldFetch = useSelector(getFetchStatus);
  const isLoading = useSelector(getLoadingStatus);

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

  const mediaTypesRef = useRef(Data);

  useEffect(() => {
    mediaTypesRef.current = Data;
  }, [Data]);

  useEffect(() => {
    if (!shouldFetch) return;

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

    async function CollectMediaCovers(type, item) {
      let title;
      let author;
      if (type === "book") {
        title = item.title;
        author = item.author;
      } else {
        title = item;
      }

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
      if (type === "book") {
        blockInfo = await grabOpenLibraryData({ title, author }); // { title, author, first_publish_year, number_of_pages }
      } else {
        blockInfo = { title };
      }

      setCollectedCoversBlocks((blocks) => [
        ...blocks,
        { type, images: imgArr, blockInfo, id: nanoid() },
      ]);
      return;
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
        dispatch(finishedFetch());
      })
      .catch(console.error);

    return () => {
      isCancelled = true;
    };
  }, [shouldFetch, dispatch]);

  //--- Setting state in parent for gathering all of the data from the media text inputs ---//

  const [searchData, setSearchData] = useState(
    Data.map((media) => ({ type: media.type, text: "" }))
  );

  //----------------------------------------------------------------------------------------//

  return (
    <>
      <div
        className="InfoForm"
        style={{
          backgroundImage: `url(${BackgroundIMG})`,
        }}
      >
        <QueryCounter />
        <img src={`${MediaCollectorTitle}`} />
        <MediaCheckboxes mediaTypes={Data} />

        <ButtonGroup
          searchData={searchData}
          setSearchData={setSearchData}
          mediaTypes={Data}
        />

        <MediaInputs mediaTypes={Data} setSearchData={setSearchData} />
      </div>

      {CollectedCoversBlocks.length > 0 && (
        <TitleBlockContainer blocks={CollectedCoversBlocks} />
      )}
      {isLoading && <p>Loading...</p>}
    </>
  );
};

export default MediaCollector;
