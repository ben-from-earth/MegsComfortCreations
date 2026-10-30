'use client';

import { useState } from 'react';
import { trpc } from 'lib/trpc/client';
import SearchIcon from '@mui/icons-material/Search';
import DownloadIcon from '@mui/icons-material/Download';
import QueryCounter from '@/components/shared/query-counter';
import MediaCheckboxes, {
  MediaVisibilityMap,
} from '@/mediacollector/media-checkboxes';
import MediaInputs from '@/mediacollector/media-inputs';
import PNGFormatPicker from '@/mediacollector/png-format-picker';
import LoadingWidget from '@/components/shared/loading-widget';
import DatabaseSaveFailureBody from '@/mediacollector/database-save-failure-body';
import TitleBlockContainer from '@/mediacollector/title-block-container';
import CollectorHeaderFields from '@/mediacollector/collector-header-fields';
import Dialog from '@/components/ui/dialog';
import { Form } from '@/components/ui/form';
import { useCollectorForm } from './collector-form/use-collector-form';
import type { CollectorFormData } from './collector-form/collector-form-schema';
import {
  useFieldArray,
  useFormContext,
  useFormState,
  useWatch,
} from 'react-hook-form';
import { toFormImages } from './collector-form/media-item-form-schema';
import { buildPNGExportImages } from './png-export-images';
import {
  buildDatabaseSaveFailureDisplayLines,
  markSuccessfulBlocksAsInDatabase,
  type DatabaseSaveFailureDisplayLine,
} from './database-save-error-display';

const sectionHeadingClassName =
  'm-0 text-[13px] font-extrabold tracking-[1px] text-[#6b645d] uppercase';
const primaryButtonClassName =
  'flex h-12.5 cursor-pointer items-center justify-center gap-2.5 rounded-xl border-0 bg-[#3d3935] px-5.5 text-base font-extrabold text-white disabled:cursor-not-allowed disabled:bg-[#e3ddd5] disabled:text-[#5f5851]';

function SearchTitlesButton({ onClick }: { onClick: () => void }) {
  const collectionList = useWatch<CollectorFormData, 'collectionList'>({
    name: 'collectionList',
  });
  let titleCount = 0;
  for (const searchList of Object.values(collectionList ?? {})) {
    titleCount += searchList.length;
  }

  return (
    <button
      type="button"
      className={`${primaryButtonClassName} mt-auto self-stretch`}
      onClick={onClick}
    >
      <SearchIcon fontSize="small" />
      Search {titleCount} {titleCount === 1 ? 'title' : 'titles'}
    </button>
  );
}

function ExportButton() {
  const collectedData = useWatch<CollectorFormData, 'collectedData'>({
    name: 'collectedData',
  });
  const blocksNeedingCover = (collectedData ?? []).filter(
    (block) => !block.isDatabase && !block.images.some((img) => img.selected),
  ).length;

  return (
    <>
      <button
        type="submit"
        className={primaryButtonClassName}
        disabled={blocksNeedingCover > 0}
      >
        <DownloadIcon fontSize="small" />
        Save &amp; download PNG
      </button>
      {blocksNeedingCover > 0 ? (
        <p className="m-0 text-xs font-bold text-[#9a3b26]">
          {blocksNeedingCover} {blocksNeedingCover === 1 ? 'item' : 'items'}{' '}
          still {blocksNeedingCover === 1 ? 'needs' : 'need'} a cover
        </p>
      ) : null}
    </>
  );
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
  const submitErrorMessage =
    isSubmitted && errors.collectedData
      ? 'Some collected items have invalid details. Check the highlighted blocks.'
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

  const handleCollectClick = async (): Promise<void> => {
    replace([]);
    setValue('pngExportBlockIds', []);
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

    const collectResult = await collectMedia(collectionList);
    if (!collectResult) {
      return;
    }
    replace(
      collectResult.blocks.map((block) => ({
        ...block,
        images: toFormImages(block.images, block.blockInfo.spineColor),
      })),
    );
    setValue('pngExportBlockIds', collectResult.pngExportBlockIds);
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
      const images = buildPNGExportImages(
        formValues.collectedData,
        formValues.pngExportBlockIds,
      );
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
    <main className="mx-auto flex max-w-440 flex-col gap-9 px-6 pt-7 pb-16">
      <form id={formId} onSubmit={handleCreatePngClick}>
        <section
          aria-label="Media Collector"
          className="overflow-hidden rounded-[20px] border border-[#e6e0d8] bg-white shadow-[0_8px_30px_rgba(60,50,40,0.07)]"
        >
          <div className="flex flex-wrap items-center justify-between gap-3 border-b border-[#e3e9de] bg-[#f3f6f0] px-7 py-4.5">
            <h1 className="m-0 font-(family-name:--font-hand) text-[52px] leading-[0.9] font-normal tracking-[1px]">
              Meg’s Media Collector
            </h1>
            <QueryCounter />
          </div>

          <div className="flex flex-wrap gap-7 px-7 pt-6 pb-7">
            <div className="flex flex-[1_1_250px] flex-col gap-3.5">
              <h2 className={sectionHeadingClassName}>1 · Order</h2>
              <CollectorHeaderFields />
            </div>

            <div className="flex flex-[2_1_440px] flex-col gap-3.5 border-l border-[#ece6de] pl-7">
              <h2 className={sectionHeadingClassName}>2 · Titles to search</h2>
              <MediaCheckboxes
                visibility={visibleMediaInputs}
                onToggle={handleToggleMediaInput}
              />
              <MediaInputs visibility={visibleMediaInputs} />
              <SearchTitlesButton
                onClick={() => {
                  void handleCollectClick();
                }}
              />
            </div>

            <div className="flex flex-[1_1_250px] flex-col gap-3.5 border-l border-[#ece6de] pl-7">
              <h2 className={sectionHeadingClassName}>3 · Export</h2>
              <PNGFormatPicker />
              <ExportButton />
              {submitErrorMessage ? (
                <p className="m-0 text-xs font-bold text-red-600">
                  {submitErrorMessage}
                </p>
              ) : null}
            </div>
          </div>
        </section>
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
    </main>
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
