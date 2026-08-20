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
import {
  useFieldArray,
  useFormContext,
  useFormState,
  type FieldErrors,
} from 'react-hook-form';
import { toFormImages } from './collector-form/mediaItemFormSchema';
import { buildPNGExportImages } from './png-export-images';
import {
  buildDatabaseSaveFailureDisplayLines,
  markSuccessfulBlocksAsInDatabase,
  type DatabaseSaveFailureDisplayLine,
} from './database-save-error-display';

const backgroundImage = '/FlowerBackground.png';
const mediaCollectorTitleImage = '/MegsMediaCollector.png';

function collectorSubmitErrorMessage(
  errors: FieldErrors<CollectorFormData>,
): string | null {
  if (typeof errors.pngFormat?.message === 'string') {
    return errors.pngFormat.message;
  }
  if (typeof errors.bookClubRepeat?.message === 'string') {
    return errors.bookClubRepeat.message;
  }
  if (errors.collectedData) {
    return 'Some collected items have invalid details. Check the highlighted blocks.';
  }
  return Object.keys(errors).length > 0
    ? 'Please fix the highlighted fields and try again.'
    : null;
}

function MediaCollectorContent({
  formId,
  onSubmit,
}: {
  formId: string;
  onSubmit: ReturnType<typeof useCollectorForm>['onSubmit'];
}) {
  const { control, getValues, handleSubmit, setValue } =
    useFormContext<CollectorFormData>();
  const { errors, isSubmitted } = useFormState({ control });
  const { fields, remove, replace } = useFieldArray({
    control,
    name: 'collectedData',
    keyName: 'fieldId',
  });
  const submitErrorMessage = isSubmitted
    ? collectorSubmitErrorMessage(errors)
    : null;
  const schemaErrorBlockIds = isSubmitted
    ? fields
        .filter((_, index) => errors.collectedData?.[index] != null)
        .map((field) => field.blockID)
    : [];

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
    replace(
      blocks.map((block) => ({
        ...block,
        images: toFormImages(block.images, block.blockInfo.spineColor),
      })),
    );
  };

  const handleCreatePngClick = (
    event: React.FormEvent<HTMLFormElement>,
  ): void => {
    event.preventDefault();
    setBlockIdsWithErrors([]);
    const itemsWithoutImages = getValues('collectedData').filter((block) => {
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

    void handleSubmit(handleExportSubmit)();
  };

  const handleExportSubmit = async (
    formValues: CollectorFormData,
  ): Promise<void> => {
    setLoadingMessage(`Adding items to database`);
    setIsSavingToDatabase(true);

    try {
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
    } finally {
      setIsSavingToDatabase(false);
    }
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
      <form
        id={formId}
        onSubmit={handleCreatePngClick}
      >
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

            <div className="flex flex-col items-center gap-2">
              <div className="flex flex-row items-center gap-4">
                <Button
                  type="button"
                  variant="primary"
                  onClick={() => {
                    handleCollectClick();
                  }}
                  label="Collect Media Covers"
                  width={175}
                  fontSize={25}
                />
                <Button
                  type="submit"
                  variant="primary"
                  label="Create PNG Export"
                  width={175}
                  fontSize={25}
                />
              </div>
              {submitErrorMessage ? (
                <p className="m-0 font-['Just_Another_Hand'] text-2xl tracking-wider text-red-600">
                  {submitErrorMessage}
                </p>
              ) : null}
            </div>

            <MediaInputs visibility={visibleMediaInputs} />
          </div>
          <PNGFormatPicker />
        </div>
      </form>
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
          blockIdsWithErrors={[
            ...new Set([...blockIdsWithErrors, ...schemaErrorBlockIds]),
          ]}
        />
      )}
    </>
  );
}

export default function MediaCollector() {
  const { form, formId, onSubmit } = useCollectorForm();
  return (
    <Form {...form}>
      <MediaCollectorContent formId={formId} onSubmit={onSubmit} />
    </Form>
  );
}
