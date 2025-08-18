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
} from "../../app/collectorSlice";

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

    let isCancelled = false;
    setCollectedCoversBlocks([]);

    const work = mediaTypesRef.current
      .filter((mediaType) => mediaType.toCollect.length > 0)
      .flatMap(({ type, toCollect }) =>
        toCollect.map((t) => ({ type, toCollectItem: t }))
      );

    const promises = work.map(async ({ type, toCollectItem }) => {
      const newBlock = await dispatch(
        collectMediaCovers({ type, toCollectItem })
      ).unwrap();
      setCollectedCoversBlocks((blocks) => [...blocks, newBlock]);
    });

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

  const handleCollectClick = () => {
    dispatch(setCollectText({ searchData }));
    dispatch(collectMedia());
    setSearchData(stateData.map((media) => ({ type: media.type, text: "" })));
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
        <MediaCheckboxes mediaTypes={stateData} />

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
