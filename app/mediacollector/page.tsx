'use client';
// image collection from assets
import backgroundImage from '@/public/FlowerBackground.png';
import MediaCollectorTitle from '@/public/MegsMediaCollector.png';

// react and redux
import { useEffect, useRef, useState } from 'react';
import { useSelector } from 'react-redux';
import { useAppDispatch } from '@/lib/state/store';

// library imports
import axios from 'axios';

// components
import Image from 'next/image';
import QueryCounter from '@/app/components/QueryCounter';
import MediaCheckboxes from '@/app/mediacollector/MediaCheckboxes';
import MediaInputs from '@/app/mediacollector/MediaInputs';
import PNGFormatPicker from '@/app/mediacollector/PNGFormatPicker';
import ButtonGroup from '@/app/mediacollector/ButtonGroup';
import LoadingWidget from '@/app/components/LoadingWidget';
import DatabaseSavedWidget from '@/app/mediacollector/DatabaseSavedWidget';
import TitleBlockContainer from '@/app/mediacollector/TitleBlockContainer';

// imports from the collector state slice
import {
  collectBlockInformation,
  collectedBlockInformation,
  finishedFetch,
  finishedLoad,
  getFetchStatus,
  getLoadingStatus,
  mediaData,
  mediaTypeDefinitions,
  setCollectList,
  startFetch,
  startLoad,
} from '@/lib/state/slices/collectorSlice';

// imports from the database state slice
import {
  clearDatabaseData,
  databaseDataPerType,
  removeFromDatabaseData,
  sendToDatabase,
} from '@/lib/state/slices/databaseDataSlice';

// imports from the png state slice
import {
  clearPNGCollectionList,
  ImageData,
  removeFromPNGCollectionList,
  selectPNGList,
} from '@/lib/state/slices/pngCollectionSlice';

//interfaces and types
import {
  databaseSaveServerResponse,
  MediaType,
} from '@/lib/interfaces/globalInterfaces';
import { titleOutputObj } from '@/lib/helpers/titleCollectionListConversion';
import { ErrorResponse } from '@/app/api/api-Errors';

export default function MediaCollector() {
  //setup connection to the redux slice and associated variables
  const dispatch = useAppDispatch();
  const stateData: mediaTypeDefinitions[] = useSelector(mediaData);
  const shouldFetch: boolean = useSelector(getFetchStatus);
  const isLoading: boolean = useSelector(getLoadingStatus);
  const pngCollectionList: ImageData[] = useSelector(selectPNGList);

  // setup states used throughout the component
  const [CollectedCoversBlocks, setCollectedCoversBlocks] = useState<
    collectedBlockInformation[]
  >([]);
  const [searchData, setSearchData] = useState<
    { type: MediaType; titleSearchList: titleOutputObj[] }[]
  >(stateData.map((media) => ({ type: media.type, titleSearchList: [] })));
  const [pngTemplateChecks, setPNGTemplateChecks] = useState<boolean[]>([
    false,
    false,
  ]);
  const [pngTemplate, setPNGTemplate] = useState<number | undefined>();
  const [pngError, setPNGError] = useState<boolean>(false);
  const [loadingMessage, setLoadingMessage] = useState<string>('');
  const [databaseSaved, setDatabaseSaved] = useState<boolean>(false);
  const [databaseSavedData, setDatabaseSavedData] =
    useState<databaseSaveServerResponse>([]);

  //refs for useEffect
  const mediaTypesRef = useRef(stateData);

  //update ref whenever stateData updates
  //stateData holds data about showing checkboxes, and the titleCollectionList data.
  useEffect(() => {
    mediaTypesRef.current = stateData;
  }, [stateData]);

  //on mount, reset query count in local storage if its a new day
  //on unmount reset the visual state
  useEffect(() => {
    const lastQueryDate = localStorage.getItem('lastQueryDate');
    const today = new Date().toISOString().split('T')[0];
    if (lastQueryDate !== today) {
      localStorage.setItem('queryCount', '0');
      localStorage.setItem('lastQueryDate', today);
    }
  }, []);

  //if should fetch comes true, set off chain of events to collect media covers
  useEffect(() => {
    if (!shouldFetch) return;

    let cancelled = false;
    setCollectedCoversBlocks([]);

    const work = mediaTypesRef.current
      .filter((m) => m.titleCollectionList.length > 0)
      .flatMap(({ type, titleCollectionList }) =>
        titleCollectionList.map((t) => ({ type, toCollectItem: t })),
      );

    // Kick off thunks
    const tasks = work.map(({ type, toCollectItem }) =>
      dispatch(collectBlockInformation({ type, toCollectItem })),
    );

    //IIFE for async promise collection and collected covers block setting
    (async () => {
      try {
        // unwrap in parallel
        const results = await Promise.allSettled(tasks.map((t) => t.unwrap()));
        if (cancelled) return;

        const blocks = results
          .filter((r) => r.status === 'fulfilled')
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
      tasks.forEach((t) => t.abort?.());
    };
  }, [shouldFetch, dispatch]);

  //Set of actions that set of Media Cover Collection
  const handleCollectClick = (): void => {
    //clear database information and PNG collection list if any information persists so blocks can populate appropriately
    dispatch(clearDatabaseData());
    dispatch(clearPNGCollectionList());

    //Set the collection list from the text areas
    dispatch(setCollectList(searchData));

    //count number of items for loading widget
    let count = 0;
    searchData.map((type) => {
      count += type.titleSearchList.length;
    });
    setLoadingMessage(`Gathering ${count} media covers`);

    //tell UI were loading and to kick off fetch in the above useEffect
    dispatch(startLoad());
    dispatch(startFetch());
  };

  //Send block information to the database and give responses to the Database Saved Widget to appropriately show successes and errors
  const handleDatabaseClick = async (
    databaseData: databaseDataPerType[],
  ): Promise<void> => {
    const responses = await dispatch(sendToDatabase(databaseData)).unwrap();
    setDatabaseSavedData(responses);
    setDatabaseSaved(true);
  };

  //Handle creation of PNG from all covers. This is the main finishing product of the app
  const handlePNGClick = async (): Promise<void> => {
    if (!pngTemplate) {
      setPNGError(true);
      return;
    }

    setLoadingMessage(`Putting together PNG export`);
    dispatch(startLoad());

    try {
      const res = await axios.post<Blob | ErrorResponse>(
        `/api/png/create`,
        { template: pngTemplate, images: pngCollectionList },
        {
          responseType: 'blob',
          // Accept both, let the server decide
          headers: { Accept: 'image/png, application/zip' },
          // increase timeout for large zips
          timeout: 120000,
        },
      );

      // Determine filename & extension
      const contentType = (res.headers['content-type'] || '').toLowerCase();
      const contentDisp = res.headers['content-disposition'] || '';

      // Try to parse a filename from Content-Disposition
      let filename = (() => {
        const m = /filename\*?=(?:UTF-8'')?["']?([^"';]+)["']?/i.exec(
          contentDisp,
        );
        return m ? decodeURIComponent(m[1]) : null;
      })();

      // Fallback filename by MIME
      if (!filename) {
        const ext = contentType.includes('zip')
          ? 'zip'
          : contentType.includes('png')
            ? 'png'
            : 'bin';
        filename = `MCC_PNG_export.${ext}`;
      }

      // Create object URL and trigger download
      if ('error' in res.data === false) {
        const blob = res.data; // already a Blob because responseType: 'blob'
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        // In case the browser blocks a.click() without it being in DOM:
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
      }
    } catch (err) {
      // Surface message if server returned JSON error but axios treated as blob
      // (axios throws for non-2xx; this is just a friendly message)
      console.error('Download failed:', err);
    } finally {
      dispatch(finishedLoad());
    }
  };

  //Delete a block action for if any end up uneeded or user has any reason to not want to add information to database or PNG export.
  const handleDeleteBlock = (
    blockID: string,
    type: MediaType,
    deleteBlock: boolean,
    urls: string[],
  ) => {
    setCollectedCoversBlocks((prev) =>
      prev.filter((block) => block.blockID !== blockID),
    );
    dispatch(removeFromDatabaseData({ blockID, type, deleteBlock }));
    for (const url of urls) {
      dispatch(removeFromPNGCollectionList({ url }));
    }
  };

  return (
    <>
      <div
        className="border-b-darkpink relative box-border flex h-fit w-full flex-col items-center border-b-5 bg-cover pt-1 shadow-[5px_5px_30px_rgba(0,0,0,0.3)]"
        style={{
          backgroundImage: `url(${backgroundImage.src})`,
        }}
      >
        <QueryCounter />
        <Image
          alt="Megs Media Collector Title"
          src={MediaCollectorTitle}
          width={576}
        />
        <MediaCheckboxes mediaTypes={stateData} setSearchData={setSearchData} />

        <ButtonGroup
          onCollect={handleCollectClick}
          onPNG={handlePNGClick}
          onDatabase={handleDatabaseClick}
        />

        <MediaInputs mediaTypes={stateData} setSearchData={setSearchData} />
        <PNGFormatPicker
          pngTemplateChecks={pngTemplateChecks}
          pngError={pngError}
          setPNGError={setPNGError}
          setPNGTemplate={setPNGTemplate}
          setPNGTemplateChecks={setPNGTemplateChecks}
        />
      </div>
      {isLoading && <LoadingWidget message={loadingMessage} />}
      {databaseSaved && (
        <DatabaseSavedWidget
          data={databaseSavedData}
          close={() => setDatabaseSaved(false)}
        />
      )}
      {CollectedCoversBlocks.length > 0 && (
        <TitleBlockContainer
          blocks={CollectedCoversBlocks}
          handleDeleteBlock={handleDeleteBlock}
        />
      )}
    </>
  );
}
