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
  collectBlockInformation,
} from "../../state/collectorSlice";
import {
  clearDatabaseData,
  removeFromDatabaseData,
} from "../../state/databaseDataSlice";
import {
  clearPNGCollectionList,
  removeFromPNGCollectionList,
  selectPNGList,
} from "../../state/pngCollectionSlice";
import LoadingWidget from "./LoadingWidget";

const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

const MediaCollector = () => {
  //setup connection to the redux slice
  const dispatch = useDispatch();
  const stateData = useSelector(mediaData);
  const shouldFetch = useSelector(getFetchStatus);
  const isLoading = useSelector(getLoadingStatus);
  const pngCollectionList = useSelector(selectPNGList);

  // setup states used throughout the component
  const [CollectedCoversBlocks, setCollectedCoversBlocks] = useState([]);
  const [searchData, setSearchData] = useState(
    stateData.map((media) => ({ type: media.type, text: "" }))
  );
  const [pngTemplateChecks, setPNGTemplateChecks] = useState([false, false]);
  const [pngTemplate, setPNGTemplate] = useState();
  const [pngError, setPNGError] = useState(false);
  const [searchCount, setSearchCount] = useState(0);

  //refs for useEffect
  const mediaTypesRef = useRef(stateData);

  //update ref whenever stateData updates
  //stateData holds data about showing checkboxes, and the toCollect data.
  useEffect(() => {
    mediaTypesRef.current = stateData;
  }, [stateData]);

  //on mount, reset query count if its a new day
  useEffect(() => {
    const lastQueryDate = localStorage.getItem("lastQueryDate");
    const today = new Date().toISOString().split("T")[0];
    if (lastQueryDate !== today) {
      localStorage.setItem("queryCount", 0);
      localStorage.setItem("lastQueryDate", today);
    }
  }, []);

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
      dispatch(collectBlockInformation({ type, toCollectItem }))
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
    dispatch(clearPNGCollectionList());
    console.log(searchData);
    dispatch(setCollectText({ searchData }));

    //count number of items for loading widget
    let count = 0;
    searchData.map((type) => {
      if (type.text.length > 0) {
        console.log(type.text.split(",").length);
        count += type.text.split(",").length;
      }
    });
    setSearchCount(count);

    dispatch(collectMedia());
  };

  const handlePNGClick = async () => {
    if (!pngTemplate) {
      setPNGError(true);
      return;
    }

    console.log(pngCollectionList);

    const res = await fetch(`${serverDomain}/print-png`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        template: pngTemplate,
        images: pngCollectionList,
      }),
    });
    if (!res.ok) {
      throw new Error(`Server Error ${res.status}`);
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "output.png";
    a.click();
  };

  const handleDeleteBlock = ({ blockID, type, deleteBlock, urls }) => {
    setCollectedCoversBlocks((prev) =>
      prev.filter((block) => block.blockID !== blockID)
    );
    dispatch(removeFromDatabaseData({ blockID, type, deleteBlock }));
    for (let url of urls) {
      dispatch(removeFromPNGCollectionList({ url }));
    }
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

        <ButtonGroup
          onCollect={handleCollectClick}
          onPNG={handlePNGClick}
          PNGButtonAllowed={pngTemplateChecks.some((val) => val === true)}
        />

        <MediaInputs mediaTypes={stateData} setSearchData={setSearchData} />
        <div className="PNGFormat">
          <p
            style={{
              visibility: pngError ? "visible" : "hidden",
              color: "red",
            }}
          >
            Please select a PNG template option
          </p>
          <div className="pngTemplateSelection">
            <label className="MCC-font">
              <input
                id={"3mm"}
                type="checkbox"
                checked={pngTemplateChecks[0]}
                onChange={(e) => {
                  if (e.target.checked === true) {
                    setPNGError(false);
                    setPNGTemplateChecks([true, false]);
                    setPNGTemplate(3);
                  } else {
                    setPNGTemplateChecks((prev) => [false, prev[1]]);
                    setPNGTemplate();
                  }
                }}
              />
              3mm PNG Format
            </label>
            <label className="MCC-font">
              <input
                id={"5mm"}
                type="checkbox"
                checked={pngTemplateChecks[1]}
                onChange={(e) => {
                  if (e.target.checked === true) {
                    setPNGTemplateChecks([false, true]);
                    setPNGTemplate(5);
                    setPNGError(false);
                  } else {
                    setPNGTemplateChecks((prev) => [prev[0], false]);
                    setPNGTemplate();
                  }
                }}
              />
              5mm PNG Format
            </label>
          </div>
        </div>
      </div>
      {isLoading && <LoadingWidget searchCount={searchCount} />}
      {CollectedCoversBlocks.length > 0 && (
        <TitleBlockContainer
          blocks={CollectedCoversBlocks}
          handleDeleteBlock={handleDeleteBlock}
        />
      )}
    </>
  );
};

export default MediaCollector;
