'use client';
// image collection from assets
import backgroundImage from 'public/FlowerBackground.png';
import MediaCollectorTitle from 'public/MegsMediaCollector.png';

// react and redux
import { useState } from 'react';

// library imports
import { trpc } from 'lib/trpc/client';

// components
import Image from 'next/image';
import QueryCounter from '@/shared/QueryCounter';
// import MediaCheckboxes from '@/mediacollector/MediaCheckboxes';
// import MediaInputs from '@/mediacollector/MediaInputs';
import PNGFormatPicker from '@/mediacollector/PNGFormatPicker';
import LoadingWidget from '@/shared/LoadingWidget';
import InformationalDialog from '@/mediacollector/InformationalDialog';
import TitleBlockContainer from '@/mediacollector/TitleBlockContainer';
import TextInput from '@/shared/TextInput';

//interfaces and types
import { DatabaseSaveServerResponse } from 'lib/interfaces/globalInterfaces';
import Button from '@/components/ui/Button';
import { useCollectorForm } from './collector-form/use-collector-form';
import type { CollectorFormData } from './collector-form/collectorFormSchema';
import { FormProvider, useFormContext } from 'react-hook-form';
import titleCollectionListConversion from 'lib/helpers/titleCollectionListConversion';

function MediaCollectorContent() {
  const { onSubmit } = useCollectorForm();

  // setup states used throughout the component
  const { watch, setValue } = useFormContext<CollectorFormData>();

  const formValues = watch();

  const [pngError, setPNGError] = useState<boolean>(false);
  const [blockIdsWithErrors, setBlockIdsWithErrors] = useState<string[]>([]);
  const [loadingMessage, setLoadingMessage] = useState<string>('');
  const [databaseSaved, setDatabaseSaved] = useState<boolean>(false);
  const [showInformationalDialog, setShowInformationalDialog] =
    useState<boolean>(false);
  const [informationalDialogText, setInformationalDialogText] =
    useState<string>('');
  const [isSavingToDatabase, setIsSavingToDatabase] = useState<boolean>(false);
  const [databaseSavedData, setDatabaseSavedData] =
    useState<DatabaseSaveServerResponse>([]);

  //refs for useEffect
  // const mediaTypesRef = useRef(stateData);

  // trpc functions
  const { mutateAsync: createPNG, isPending: isCreatingPNG } =
    trpc.png.create.useMutation();
  const utils = trpc.useUtils();
  const { mutateAsync: collectMedia, isPending: isCollectingMedia } =
    trpc.collect.collectMedia.useMutation({
      onSuccess: () => {
        const date = new Date().toLocaleDateString('en-CA', {
          timeZone: 'America/New_York',
        });
        // Invalidate today's query count so QueryCounter refetches
        void utils.database.getQueryCount.invalidate({ date });
      },
    });

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

  const handlePNGClick = async (): Promise<void> => {
    setBlockIdsWithErrors([]);
    const itemsWithoutImages = formValues.collectedData.filter((block) => {
      const selectedImages = block.images.filter((img) => img.selected);
      if (block.isDatabase) {
        return false;
      }
      return selectedImages.length === 0;
    });
    const numWithoutImages = itemsWithoutImages.length;
    if (numWithoutImages > 0) {
      setInformationalDialogText(
        `There are ${numWithoutImages} blocks without selected images that will not be saved to the database. Please select at least one image per block or delete the block.`,
      );
      setBlockIdsWithErrors(itemsWithoutImages.map((block) => block.blockID));
      setShowInformationalDialog(true);
      return;
    }

    if (!formValues.pngFormat) {
      setPNGError(true);
      return;
    }

    if (formValues.bookClubRepeat < 1) {
      setInformationalDialogText(`Book Club Repeat Number must be at least 1.`);
      setShowInformationalDialog(true);
      return;

      setLoadingMessage(`Adding items to database`);
      setIsSavingToDatabase(true);
    }

    const result = await onSubmit(formValues);

    const someHaveDatabaseErrors = result.some((res) => 'error' in res);

    if (someHaveDatabaseErrors) {
      setInformationalDialogText(
        `There were database errors when saving some of the media blocks. Check status of database`,
      );
      setShowInformationalDialog(true);
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
      repeatCount: formValues.bookClubRepeat,
    });
    setIsSavingToDatabase(false);

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

  const handleDeleteBlock = (blockID: string) => {
    setValue(
      'collectedData',
      formValues.collectedData.filter((block) => block.blockID !== blockID),
    );
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
          <TextInput
            onChange={(e) => {
              setValue('bookClubRepeat', Number(e.target.value));
            }}
            label={'Book Club Repeat Number'}
            variant="normal"
            value={formValues.bookClubRepeat.toString()}
          />

          {/* <MediaCheckboxes mediaTypes={stateData} /> */}

          <div className="flex flex-row items-center gap-4">
            <Button
              variant="primary"
              onClick={() => {
                handleCollectClick();
              }}
              label={'Collect Media Covers'}
              width={175}
              fontSize={25}
            />
            <Button
              variant="primary"
              onClick={() => handlePNGClick()}
              label={'Create PNG Export'}
              width={175}
              fontSize={25}
            />
          </div>

          <TextInput
            variant="multiline"
            label={`Book Titles`}
            rows={5}
            onChange={(e) => {
              const titleSearchList = titleCollectionListConversion(
                e.target.value,
              );
              setValue(`collectionList.book`, titleSearchList);
            }}
          />
          <TextInput
            variant="multiline"
            label={`Movie Titles`}
            rows={5}
            onChange={(e) => {
              const titleSearchList = titleCollectionListConversion(
                e.target.value,
              );
              setValue(`collectionList.movie`, titleSearchList);
            }}
          />
          <TextInput
            variant="multiline"
            label={`Video Game Titles`}
            rows={5}
            onChange={(e) => {
              const titleSearchList = titleCollectionListConversion(
                e.target.value,
              );
              setValue(`collectionList.videoGame`, titleSearchList);
            }}
          />
          <TextInput
            variant="multiline"
            label={`Album Titles`}
            rows={5}
            onChange={(e) => {
              const titleSearchList = titleCollectionListConversion(
                e.target.value,
              );
              setValue(`collectionList.album`, titleSearchList);
            }}
          />
        </div>
        <PNGFormatPicker pngError={pngError} setPNGError={setPNGError} />
      </div>
      {(isCollectingMedia || isCreatingPNG || isSavingToDatabase) && (
        <LoadingWidget message={loadingMessage} />
      )}
      {databaseSaved && (
        <InformationalDialog
          variant="databaseSave"
          data={databaseSavedData}
          close={() => setDatabaseSaved(false)}
        />
      )}
      {showInformationalDialog && (
        <InformationalDialog
          variant="informationalOnly"
          infoText={informationalDialogText}
          close={() => setShowInformationalDialog(false)}
        />
      )}
      {formValues.collectedData.length > 0 && (
        <TitleBlockContainer
          handleDeleteBlock={handleDeleteBlock}
          blockIdsWithErrors={blockIdsWithErrors}
        />
      )}
    </>
  );
}

export default function MediaCollector() {
  const { form } = useCollectorForm();
  return (
    <FormProvider {...form}>
      <MediaCollectorContent />
    </FormProvider>
  );
}
