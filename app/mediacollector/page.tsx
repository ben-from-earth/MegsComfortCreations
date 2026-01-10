'use client';
// image collection from assets
import backgroundImage from 'public/FlowerBackground.png';
import MediaCollectorTitle from 'public/MegsMediaCollector.png';

// react and redux
import { useEffect, useState } from 'react';

// library imports
import { trpc } from 'lib/trpc/client';

// components
import Image from 'next/image';
import QueryCounter from '@/shared/QueryCounter';
// import MediaCheckboxes from '@/mediacollector/MediaCheckboxes';
import MediaInputs from '@/mediacollector/MediaInputs';
import PNGFormatPicker from '@/mediacollector/PNGFormatPicker';
import LoadingWidget from '@/shared/LoadingWidget';
import DatabaseSavedWidget from '@/mediacollector/DatabaseSavedWidget';
import TitleBlockContainer from '@/mediacollector/TitleBlockContainer';
import TextInput from '@/shared/TextInput';

//interfaces and types
import { DatabaseSaveServerResponse } from 'lib/interfaces/globalInterfaces';
import Button from '@/shared/Button';
import { useCollectorForm } from './collector-form/use-collector-form';
import type { CollectorFormData } from './collector-form/collectorFormSchema';
import { FormProvider, useFormContext } from 'react-hook-form';

function MediaCollectorContent() {
  const { onSubmit } = useCollectorForm();

  // setup states used throughout the component
  const { watch, setValue } = useFormContext<CollectorFormData>();

  const formValues = watch();
  console.log('Form Values:', formValues);

  const [pngError, setPNGError] = useState<boolean>(false);
  const [loadingMessage, setLoadingMessage] = useState<string>('');
  const [databaseSaved, setDatabaseSaved] = useState<boolean>(false);
  const [databaseSavedData, setDatabaseSavedData] =
    useState<DatabaseSaveServerResponse>([]);

  //refs for useEffect
  // const mediaTypesRef = useRef(stateData);

  // trpc functions
  const { mutateAsync: createPNG, isPending: isCreatingPNG } =
    trpc.png.create.useMutation();
  const { mutateAsync: collectMedia, isPending: isCollectingMedia } =
    trpc.collect.collectMedia.useMutation();

  //update ref whenever stateData updates
  //stateData holds data about showing checkboxes, and the titleCollectionList data.
  // useEffect(() => {
  //   mediaTypesRef.current = stateData;
  // }, [stateData]);

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

  const handleDatabaseClick = async () => {
    const result = await onSubmit(formValues);
    setDatabaseSaved(true);
    setDatabaseSavedData(result);
  };

  //Set of actions that set of Media Cover Collection
  const handleCollectClick = async (): Promise<void> => {
    setValue('collectedData', []);
    let count = 0;
    for (const searchList of Object.values(formValues.collectionList)) {
      count += searchList.length;
    }

    if (count === 0) {
      alert('No media titles to collect. Please add titles first.');
      return;
    }

    setLoadingMessage(`Gathering ${count} media covers`);

    const blocks = await collectMedia(formValues.collectionList);
    if (!blocks) {
      return;
    } else {
      setValue('collectedData', blocks);
    }
  };

  //Handle creation of PNG from all covers. This is the main finishing product of the app

  const handlePNGClick = async (): Promise<void> => {
    if (!formValues.pngFormat) {
      setPNGError(true);
      return;
    }

    setLoadingMessage(`Putting together PNG export`);
    const images = formValues.collectedData.map((block) => {
      let keptImages: { url: string; selected: boolean }[] = [];
      if (block.isDatabase) {
        keptImages = block.images;
      } else {
        keptImages = block.images.filter((img) => img.selected);
      }

      const url = keptImages.map((img) => img.url)[0];
      const type = block.type;
      const spineColor = block.blockInfo.spineColor;
      return {
        url,
        type,
        spineColor,
      };
    });

    const { mime, filename, dataBase64 } = await createPNG({
      template: Number(formValues.pngFormat) as 3 | 5,
      images,
      customerName: formValues.customerName,
      orderNumber: formValues.orderNumber,
    });

    const byteArray = Uint8Array.from(atob(dataBase64), (c) => c.charCodeAt(0));
    const blob = new Blob([byteArray], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  };

  //Delete a block action for if any end up uneeded or user has any reason to not want to add information to database or PNG export.
  const handleDeleteBlock = (blockID: string) => {
    setValue(
      'collectedData',
      formValues.collectedData.filter((block) => block.blockID !== blockID),
    );
  };

  // const databaseData: DatabaseDataPerType[] = useSelector(selectDatabaseData);

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
        <div className="mt-2 flex flex-col items-center gap-4">
          <TextInput
            onChange={(e) => {
              setValue('customerName', e.target.value);
            }}
            label={'Customer Full Name'}
            variant="normal"
            value={formValues.customerName}
          />
          <TextInput
            onChange={(e) => {
              setValue('orderNumber', e.target.value);
            }}
            label={'Order Number'}
            variant="normal"
            value={formValues.orderNumber}
          />

          {/* <MediaCheckboxes mediaTypes={stateData} /> */}

          <div className="flex flex-row items-center gap-4">
            <Button
              onClick={() => {
                handleCollectClick();
              }}
              label={'Collect Media Covers'}
              width={175}
              fontSize={25}
            />
            <Button
              onClick={() => handleDatabaseClick()}
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

          <MediaInputs />
        </div>
        <PNGFormatPicker pngError={pngError} setPNGError={setPNGError} />
      </div>
      {(isCollectingMedia || isCreatingPNG) && (
        <LoadingWidget message={loadingMessage} />
      )}
      {databaseSaved && (
        <DatabaseSavedWidget
          data={databaseSavedData}
          close={() => setDatabaseSaved(false)}
        />
      )}
      {formValues.collectedData.length > 0 && (
        <TitleBlockContainer handleDeleteBlock={handleDeleteBlock} />
      )}
    </>
  );
}

export default function MediaCollector() {
  const { form } = useCollectorForm();
  return (
    <FormProvider {...form}>
      {/* Optionally wrap in a <form> for native submit */}
      {/* <form onSubmit={form.handleSubmit(onSubmit)}> */}
      <MediaCollectorContent />
      {/* </form> */}
    </FormProvider>
  );
}
