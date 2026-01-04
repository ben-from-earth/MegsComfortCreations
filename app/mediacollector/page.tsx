'use client';
// image collection from assets
import backgroundImage from 'public/FlowerBackground.png';
import MediaCollectorTitle from 'public/MegsMediaCollector.png';

// react and redux
import { useEffect, useRef, useState } from 'react';
import { useSelector } from 'react-redux';
import { useAppDispatch } from 'lib/state/store';

// library imports
import { trpc } from 'lib/trpc/client';

// components
import Image from 'next/image';
import QueryCounter from '@/shared/QueryCounter';
import MediaCheckboxes from '@//mediacollector/MediaCheckboxes';
import MediaInputs from '@//mediacollector/MediaInputs';
import PNGFormatPicker from '@//mediacollector/PNGFormatPicker';
import LoadingWidget from '@/shared/LoadingWidget';
import DatabaseSavedWidget from '@//mediacollector/DatabaseSavedWidget';
import TitleBlockContainer from '@//mediacollector/TitleBlockContainer';
import TextInput from '@/shared/TextInput';

// imports from the collector state slice
import {
  collectBlockInformation,
  CollectedBlockInformation,
  finishedFetch,
  finishedLoad,
  getFetchStatus,
  getLoadingStatus,
  mediaData,
  mediaTypeDefinitions,
  setCollectList,
  startFetch,
  startLoad,
} from 'lib/state/slices/collectorSlice';

// imports from the database state slice
import {
  clearDatabaseData,
  DatabaseDataPerType,
  removeFromDatabaseData,
  selectDatabaseData,
} from 'lib/state/slices/databaseDataSlice';

// imports from the png state slice
import {
  clearPNGCollectionList,
  ImageData,
  removeFromPNGCollectionList,
  selectPNGList,
} from 'lib/state/slices/pngCollectionSlice';

//interfaces and types
import {
  DatabaseSaveServerResponse,
  MediaType,
  BookInsert,
  MovieInsert,
  VideoGameInsert,
  AlbumInsert,
} from 'lib/interfaces/globalInterfaces';
import { titleOutputObj } from 'lib/helpers/titleCollectionListConversion';
import Button from '@/shared/Button';
import { useCollectorForm } from './collector-form/use-collector-form';
import type { CollectorFormData } from './collector-form/collectorFormSchema';
import { FormProvider, useFormContext } from 'react-hook-form';

function MediaCollectorContent({
  onSubmit,
}: {
  onSubmit: (data: CollectorFormData) => void;
}) {
  //setup connection to the redux slice and associated variables
  const dispatch = useAppDispatch();
  const stateData: mediaTypeDefinitions[] = useSelector(mediaData);
  const shouldFetch: boolean = useSelector(getFetchStatus);
  const isLoading: boolean = useSelector(getLoadingStatus);
  const pngCollectionList: ImageData[] = useSelector(selectPNGList);

  // setup states used throughout the component
  const { watch, setValue } = useFormContext<CollectorFormData>();

  // ---------
  // these states will get integrated into the form. collected covers blocks will be derived from form data
  const [CollectedCoversBlocks, setCollectedCoversBlocks] = useState<
    CollectedBlockInformation[]
  >([]);
  const [searchData, setSearchData] = useState<
    { type: MediaType; titleSearchList: titleOutputObj[] }[]
  >(stateData.map((media) => ({ type: media.type, titleSearchList: [] })));

  // ---------
  const orderNumber = watch('orderNumber');
  const customerName = watch('customerName');

  // onSubmit comes from parent wrapper that owns the form instance

  const [pngTemplateChecks, setPNGTemplateChecks] = useState<boolean[]>([
    false,
    false,
  ]);
  const [pngTemplate, setPNGTemplate] = useState<number | undefined>();
  const [pngError, setPNGError] = useState<boolean>(false);
  const [loadingMessage, setLoadingMessage] = useState<string>('');
  const [databaseSaved, setDatabaseSaved] = useState<boolean>(false);
  const [databaseSavedData, setDatabaseSavedData] =
    useState<DatabaseSaveServerResponse>([]);

  //refs for useEffect
  const mediaTypesRef = useRef(stateData);

  // trpc functions
  const { mutateAsync: databaseSave } = trpc.database.save.useMutation();
  const { mutateAsync: linkGenres } = trpc.genres.link.useMutation();
  const { mutateAsync: createPNG } = trpc.png.create.useMutation();

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
    databaseData: DatabaseDataPerType[],
  ): Promise<void> => {
    const responses: DatabaseSaveServerResponse = [];
    for (const media of databaseData) {
      if (media.type === 'book') {
        for (const book of media.data) {
          const title = book.title; // titleRearrange already applied earlier upstream
          const bookData = {
            title,
            author: book.author,
            pageCount: book.pageCount,
            pubYear: book.pubYear,
            imageUrls: book.images.map((item) => item.src),
            spineColor: book.spineColor,
            genres: book.genres,
            blockID: book.blockID,
          };
          try {
            const res = await databaseSave({
              type: 'book',
              mediaData: bookData as BookInsert,
            });
            if ('error' in res) {
              responses.push(res);
            } else {
              const bookDatabaseID = res.actionAttemptItem.id;
              const linkRes = await linkGenres({
                bookID: bookDatabaseID,
                genres: book.genres ?? [],
              });
              responses.push({ ...res, ...linkRes });
            }
          } catch {
            responses.push({
              actionAttemptItem: bookData,
              type: 'book',
              errors: [
                'Server Error during save',
                `${bookData.title} did not save to the database`,
              ],
              error: 'Server Error',
              message: `There was a server error during save attempt for ${bookData.title}`,
            });
          }
        }
      } else {
        for (const other of media.data) {
          const title = other.title;
          const otherData: MovieInsert | VideoGameInsert | AlbumInsert = {
            title,
            imageUrls: other.images.map((img) => img.src),
            spineColor: other.spineColor,
            blockID: other.blockID,
          };
          try {
            const res = await databaseSave({
              type: media.type,
              mediaData: otherData,
            });
            responses.push(res);
          } catch {
            responses.push({
              actionAttemptItem: otherData,
              type: media.type,
              errors: [
                'Server Error during save',
                `${otherData.title} did not save to the database`,
              ],
              error: 'Server Error',
              message: `There was a server error during save attempt for ${otherData.title}`,
            });
          }
        }
      }
    }
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
      const res = await createPNG({
        template: pngTemplate as number as 3 | 5,
        images: pngCollectionList,
      });
      const { mime, filename, dataBase64 } = res as {
        mime: string;
        filename: string;
        dataBase64: string;
      };
      const byteArray = Uint8Array.from(atob(dataBase64), (c) =>
        c.charCodeAt(0),
      );
      const blob = new Blob([byteArray], { type: mime });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
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

  const databaseData: DatabaseDataPerType[] = useSelector(selectDatabaseData);

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
        <div className="mt-2 flex flex-col content-center gap-4">
          <TextInput
            onChange={(e) => {
              setValue('customerName', e.target.value);
            }}
            label={'Customer'}
            variant="normal"
            value={customerName}
          />
          <TextInput
            onChange={(e) => {
              setValue('orderNumber', e.target.value);
            }}
            label={'Order Number'}
            variant="normal"
            value={orderNumber}
          />
        </div>
        <MediaCheckboxes mediaTypes={stateData} setSearchData={setSearchData} />

        <div className="flex flex-row content-center gap-4">
          <Button
            onClick={() => {
              handleCollectClick();
            }}
            label={'Collect Media Covers'}
            width={175}
            fontSize={25}
          />
          <Button
            onClick={() => handleDatabaseClick(databaseData)}
            label={'Send to Database'}
            width={175}
            fontSize={25}
          />
          <Button
            onClick={() => handlePNGClick()}
            label={'Get PNG'}
            width={175}
            fontSize={25}
          />
        </div>

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

export default function MediaCollector() {
  const { form, onSubmit } = useCollectorForm();
  return (
    <FormProvider {...form}>
      {/* Optionally wrap in a <form> for native submit */}
      {/* <form onSubmit={form.handleSubmit(onSubmit)}> */}
      <MediaCollectorContent onSubmit={onSubmit} />
      {/* </form> */}
    </FormProvider>
  );
}
