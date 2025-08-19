// image collection from assets
import BackgroundIMG from "../../assets/FlowerBackground.png";
import MediaCollectorTitle from "../../assets/MegsMediaCollector.png";

// react and redux
import { useEffect, useRef, useState } from "react";
import { useDispatch, useSelector } from "react-redux";

//styles
import "./MediaCollector.css";

// necessary components
import MediaInputs from "./MediaInputs";
import MediaCheckboxes from "./MediaCheckboxes";
import ButtonGroup from "./ButtonGroup";
import TitleBlockContainer from "./TitleBlockContainer";
import QueryCounter from "./QueryCounter";

// necessary imports from the collector slice
import {
  collectMedia,
  finishedFetch,
  getFetchStatus,
  getLoadingStatus,
  mediaData,
  setCollectText,
  collectMediaCovers,
} from "../../state/collectorSlice";
import { clearDatabaseData } from "../../state/databaseDataSlice";

const MediaCollector = () => {
  //setup connection to the redux slice
  const dispatch = useDispatch();
  const stateData = useSelector(mediaData);
  const shouldFetch = useSelector(getFetchStatus);
  const isLoading = useSelector(getLoadingStatus);

  // setup states used throughout the component
  const [CollectedCoversBlocks, setCollectedCoversBlocks] = useState([]);
  const [searchData, setSearchData] = useState(
    stateData.map((media) => ({ type: media.type, text: "" }))
  );

  //refs for useEffect
  const mediaTypesRef = useRef(stateData);

  //update ref whenever stateData updates
  //stateData holds data about showing checkboxes, and the toCollect data.
  useEffect(() => {
    mediaTypesRef.current = stateData;
  }, [stateData]);

  useEffect(() => {
    if (!shouldFetch) return;

    let cancelled = false;
    setCollectedCoversBlocks([]);

    const work = mediaTypesRef.current
      .filter((m) => m.toCollect.length > 0)
      .flatMap(({ type, toCollect }) =>
        toCollect.map((t) => ({ type, toCollectItem: t }))
      );

    // Kick off thunks
    const tasks = work.map(({ type, toCollectItem }) =>
      dispatch(collectMediaCovers({ type, toCollectItem }))
    );

    //IIFE for async promise collection and collected covers block setting
    (async () => {
      try {
        // unwrap in parallel
        const results = await Promise.allSettled(tasks.map((t) => t.unwrap()));
        if (cancelled) return;

        const blocks = results
          .filter((r) => r.status === "fulfilled")
          .map((r) => r.value);
        setCollectedCoversBlocks(blocks);
      } catch (e) {
        console.error(e);
      } finally {
        if (!cancelled) dispatch(finishedFetch());
      }
    })();

    // on unmount/refresh: cancel thunks + mark cancelled
    return () => {
      cancelled = true;
      tasks.forEach((t) => t.abort?.()); // RTK thunks support abort()
    };
  }, [shouldFetch, dispatch]);

  const handleCollectClick = () => {
    dispatch(clearDatabaseData());
    dispatch(setCollectText({ searchData }));
    dispatch(collectMedia());
  };

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
        <MediaCheckboxes mediaTypes={stateData} setSearchData={setSearchData} />

        <ButtonGroup onCollect={handleCollectClick} />

        <MediaInputs mediaTypes={stateData} setSearchData={setSearchData} />
      </div>

      {CollectedCoversBlocks.length > 0 && (
        <TitleBlockContainer blocks={CollectedCoversBlocks} />
      )}
      {isLoading && <p>Loading...</p>}
    </>
  );
};

export default MediaCollector;
