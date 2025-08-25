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
import {
  clearDatabaseData,
  selectDatabaseData,
} from "../../state/databaseDataSlice";

const MediaCollector = () => {
  //setup connection to the redux slice
  const dispatch = useDispatch();
  const stateData = useSelector(mediaData);
  const shouldFetch = useSelector(getFetchStatus);
  const isLoading = useSelector(getLoadingStatus);
  const databaseData = useSelector(selectDatabaseData);

  // setup states used throughout the component
  const [CollectedCoversBlocks, setCollectedCoversBlocks] = useState([]);
  const [searchData, setSearchData] = useState(
    stateData.map((media) => ({ type: media.type, text: "" }))
  );
  const [pngTemplateChecks, setPngTemplateChecks] = useState([false, false]);
  const [pngTemplate, setPngTemplate] = useState();

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

  const handlePNGClick = async () => {
    if (pngTemplate && pngTemplate === 0) {
      return;
    }
    const pngImages = [];
    databaseData.map((type) => {
      for (let item of type.data) {
        let imageObjs = item.images.map((img) => ({
          url: img.src,
          type: type.type,
          spineColor: item.spineColor,
        }));
        pngImages.push(...imageObjs);
      }
    });

    const res = await fetch("http://localhost:3001/print-png", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ template: pngTemplate, images: pngImages }),
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

        <ButtonGroup onCollect={handleCollectClick} onPNG={handlePNGClick} />

        <MediaInputs mediaTypes={stateData} setSearchData={setSearchData} />
        <div className="PNGFormat">
          <p
            style={{
              visibility: pngTemplate === 0 ? "visible" : "hidden",
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
                    setPngTemplateChecks([true, false]);
                    setPngTemplate(3);
                  } else {
                    setPngTemplateChecks((prev) => [false, prev[1]]);
                    setPngTemplate(0);
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
                    setPngTemplateChecks([false, true]);
                    setPngTemplate(5);
                  } else {
                    setPngTemplateChecks((prev) => [prev[0], false]);
                    setPngTemplate(0);
                  }
                }}
              />
              5mm PNG Format
            </label>
          </div>
        </div>
      </div>

      {CollectedCoversBlocks.length > 0 && (
        <TitleBlockContainer blocks={CollectedCoversBlocks} />
      )}
      {isLoading && <p>Loading...</p>}
    </>
  );
};

export default MediaCollector;
