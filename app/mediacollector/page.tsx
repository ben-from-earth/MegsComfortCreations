'use client';

// react and redux
import { useState } from 'react';

// library imports
import { trpc } from 'lib/trpc/client';

// components
import Image from 'next/image';
import QueryCounter from '@/shared/QueryCounter';
import MediaCheckboxes, {
  MediaVisibilityMap,
} from '@/mediacollector/MediaCheckboxes';
import MediaInputs from '@/mediacollector/MediaInputs';
import PNGFormatPicker from '@/mediacollector/PNGFormatPicker';
import LoadingWidget from '@/shared/LoadingWidget';
import DatabaseSaveFailureBody from '@/mediacollector/database-save-failure-body';
import TitleBlockContainer from '@/mediacollector/TitleBlockContainer';
import TextInput from '@/shared/TextInput';

//interfaces and types
import Button from '@/components/ui/Button';
import Dialog from '@/components/ui/Dialog';
import { useCollectorForm } from './collector-form/use-collector-form';
import type { CollectorFormData } from './collector-form/collectorFormSchema';
import { FormProvider, useFormContext } from 'react-hook-form';
import { buildPNGExportImages } from './png-export-images';
import {
  buildDatabaseSaveFailureDisplayLines,
  markSuccessfulBlocksAsInDatabase,
  type DatabaseSaveFailureDisplayLine,
} from './database-save-error-display';

const backgroundImage = '/FlowerBackground.png';
const mediaCollectorTitleImage = '/MegsMediaCollector.png';

function MediaCollectorContent() {
  const { onSubmit } = useCollectorForm();

  // setup states used throughout the component
  const { watch, setValue } = useFormContext<CollectorFormData>();

  const formValues = watch();

  const [pngError, setPNGError] = useState<boolean>(false);
  const [bookClubRepeatError, setBookClubRepeatError] =
    useState<boolean>(false);
  const [blockIdsWithErrors, setBlockIdsWithErrors] = useState<string[]>([]);
  const [loadingMessage, setLoadingMessage] = useState<string>('');
  const [databaseSaved, setDatabaseSaved] = useState<boolean>(false);
  const [showInformationalDialog, setShowInformationalDialog] =
    useState<boolean>(false);
  const [informationalDialogText, setInformationalDialogText] =
    useState<string>('');
  const [isSavingToDatabase, setIsSavingToDatabase] = useState<boolean>(false);
  const [databaseSaveFailureLines, setDatabaseSaveFailureLines] = useState<
    DatabaseSaveFailureDisplayLine[]
  >([]);
  const [visibleMediaInputs, setVisibleMediaInputs] =
    useState<MediaVisibilityMap>({
      book: true,
      movie: false,
      videoGame: false,
      album: false,
    });

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
      setBookClubRepeatError(true);
      return;
    }

    // TEMPORARY: force database-save error dialog on every Create PNG attempt for UI work.
    // Remove this block when finished.
    const FORCE_PNG_CREATE_DATABASE_SAVE_ERROR_DIALOG = true;
    if (FORCE_PNG_CREATE_DATABASE_SAVE_ERROR_DIALOG) {
      const sampleBlocks = formValues.collectedData.slice(0, 2);
      const forcedFailureLines: DatabaseSaveFailureDisplayLine[] =
        sampleBlocks.length > 0
          ? sampleBlocks.map((block, index) => ({
              blockID: block.blockID,
              title: block.blockInfo.title || `Sample Title ${index + 1}`,
              blockNumber:
                formValues.collectedData.findIndex(
                  (collectedBlock) => collectedBlock.blockID === block.blockID,
                ) + 1,
              reason:
                index === 0
                  ? 'The cover image could not be saved, so this item was not added to the database.'
                  : 'A selected genre is not available in the database.',
            }))
          : [
              {
                blockID: 'forced-error-block-1',
                title: 'Forced Error Title',
                blockNumber: 1,
                reason:
                  'The cover image could not be saved, so this item was not added to the database.',
              },
              {
                blockID: 'forced-error-block-2',
                title: 'Another Forced Error',
                blockNumber: 2,
                reason: 'A selected genre is not available in the database.',
              },
            ];

      setDatabaseSaveFailureLines(forcedFailureLines);
      setBlockIdsWithErrors(
        forcedFailureLines
          .map((line) => line.blockID)
          .filter((blockID) => blockID.length > 0),
      );
      setDatabaseSaved(true);
      return;
    }

    setLoadingMessage(`Adding items to database`);
    setIsSavingToDatabase(true);

    const result = await onSubmit(formValues);

    const someHaveDatabaseErrors = result.some((res) => !res.success);

    if (someHaveDatabaseErrors) {
      const updatedCollectedData = markSuccessfulBlocksAsInDatabase(
        formValues.collectedData,
        result,
      );
      setValue('collectedData', updatedCollectedData);

      const failureLines = buildDatabaseSaveFailureDisplayLines(
        result,
        updatedCollectedData,
      );
      setDatabaseSaveFailureLines(failureLines);
      setBlockIdsWithErrors(
        failureLines
          .map((line) => line.blockID)
          .filter((blockID) => blockID.length > 0),
      );
      setDatabaseSaved(true);
      setIsSavingToDatabase(false);
      return;
    }

    setLoadingMessage(`Putting together PNG export`);
    const images = buildPNGExportImages(formValues.collectedData);

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

  const handleToggleMediaInput = (mediaType: keyof MediaVisibilityMap) => {
    setVisibleMediaInputs((prev) => {
      const next = { ...prev, [mediaType]: !prev[mediaType] };
      if (!next[mediaType]) {
        setValue(`collectionList.${mediaType}`, []);
      }
      return next;
    });
  };

  return (
    <>
      <div
        className="border-b-darkpink relative box-border flex h-fit w-full flex-col items-center border-b-5 bg-cover pt-1 shadow-[5px_5px_30px_rgba(0,0,0,0.3)]"
        style={{
          backgroundImage: `url(${backgroundImage})`,
        }}
      >
        <QueryCounter />
        <Image
          alt="Megs Media Collector Title"
          src={mediaCollectorTitleImage}
          width={576}
          height={128}
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
          <div className="flex flex-col items-center">
            <p
              className='m-0 font-["Just_Another_Hand"] text-2xl tracking-wider'
              style={{
                visibility: bookClubRepeatError ? 'visible' : 'hidden',
                color: 'red',
              }}
            >
              Book Club Repeat Number must be at least 1.
            </p>
            <TextInput
              onChange={(e) => {
                const nextValue = Number(e.target.value);
                setValue('bookClubRepeat', nextValue);
                if (nextValue >= 1) {
                  setBookClubRepeatError(false);
                }
              }}
              label={'Book Club Repeat Number'}
              variant="normal"
              value={formValues.bookClubRepeat.toString()}
            />
          </div>

          <MediaCheckboxes
            visibility={visibleMediaInputs}
            onToggle={handleToggleMediaInput}
          />

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

          <MediaInputs visibility={visibleMediaInputs} />
        </div>
        <PNGFormatPicker pngError={pngError} setPNGError={setPNGError} />
      </div>
      {(isCollectingMedia || isCreatingPNG || isSavingToDatabase) && (
        <LoadingWidget message={loadingMessage} />
      )}
      {databaseSaved && (
        <Dialog title="Error" onClose={() => setDatabaseSaved(false)}>
          <DatabaseSaveFailureBody failureLines={databaseSaveFailureLines} />
        </Dialog>
      )}
      {showInformationalDialog && (
        <Dialog
          title="Error"
          onClose={() => setShowInformationalDialog(false)}
          className="w-5/12"
        >
          <p>{informationalDialogText}</p>
        </Dialog>
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
