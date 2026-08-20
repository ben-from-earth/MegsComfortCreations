'use client';

// react and redux
import { useState } from 'react';

// library imports
import { trpc } from 'lib/trpc/client';

// components
import Image from 'next/image';
import QueryCounter from '@/components/shared/QueryCounter';
import MediaCheckboxes, {
  MediaVisibilityMap,
} from '@/mediacollector/MediaCheckboxes';
import MediaInputs from '@/mediacollector/MediaInputs';
import PNGFormatPicker from '@/mediacollector/PNGFormatPicker';
import LoadingWidget from '@/components/shared/LoadingWidget';
import DatabaseSaveFailureBody from '@/mediacollector/database-save-failure-body';
import TitleBlockContainer from '@/mediacollector/TitleBlockContainer';
import CollectorHeaderFields from '@/mediacollector/CollectorHeaderFields';

//interfaces and types
import Button from '@/components/ui/Button';
import Dialog from '@/components/ui/Dialog';
import { Form } from '@/components/ui/form';
import { useCollectorForm } from './collector-form/use-collector-form';
import type { CollectorFormData } from './collector-form/collectorFormSchema';
import { useFieldArray, useFormContext } from 'react-hook-form';
import { buildPNGExportImages } from './png-export-images';
import {
  buildDatabaseSaveFailureDisplayLines,
  markSuccessfulBlocksAsInDatabase,
  type DatabaseSaveFailureDisplayLine,
} from './database-save-error-display';

const backgroundImage = '/FlowerBackground.png';
const mediaCollectorTitleImage = '/MegsMediaCollector.png';

function MediaCollectorContent({
  onSubmit,
}: {
  onSubmit: ReturnType<typeof useCollectorForm>['onSubmit'];
}) {
  const { control, getValues, setValue, trigger } =
    useFormContext<CollectorFormData>();
  const { fields, remove, replace } = useFieldArray({
    control,
    name: 'collectedData',
    keyName: 'fieldId',
  });

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
    replace([]);
    const collectionList = getValues('collectionList');
    let count = 0;
    for (const searchList of Object.values(collectionList)) {
      count += searchList.length;
    }

    if (count === 0) {
      alert('No media titles to collect. Please add titles first.');
      return;
    }

    setLoadingMessage(`Gathering ${count} media covers`);

    const blocks = await collectMedia(collectionList);
    if (!blocks) {
      return;
    }
    replace(blocks);
  };

  const handlePNGClick = async (): Promise<void> => {
    setBlockIdsWithErrors([]);
    const formValues = getValues();
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

    const fieldsAreValid = await trigger(['pngFormat', 'bookClubRepeat']);
    if (!fieldsAreValid) {
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
      replace(updatedCollectedData);

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
    const template = formValues.pngFormat === '5' ? 5 : 3;

    const { mime, filename, dataBase64 } = await createPNG({
      template,
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
          <CollectorHeaderFields />

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
        <PNGFormatPicker />
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
      {fields.length > 0 && (
        <TitleBlockContainer
          fields={fields}
          onDelete={remove}
          blockIdsWithErrors={blockIdsWithErrors}
        />
      )}
    </>
  );
}

export default function MediaCollector() {
  const { form, onSubmit } = useCollectorForm();
  return (
    <Form {...form}>
      <MediaCollectorContent onSubmit={onSubmit} />
    </Form>
  );
}
