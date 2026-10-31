'use client';

import { use, useMemo, useState } from 'react';
import { trpc } from 'lib/trpc/client';
import SearchIcon from '@mui/icons-material/Search';
import DownloadIcon from '@mui/icons-material/Download';
import QueryCounter from '@/components/shared/query-counter';
import MediaCheckboxes, {
  MediaVisibilityMap,
} from '@/mediacollector/media-checkboxes';
import MediaInputs from '@/mediacollector/media-inputs';
import PNGFormatPicker from '@/mediacollector/png-format-picker';
import DatabaseSaveFailureBody from '@/mediacollector/database-save-failure-body';
import TitleBlockContainer, {
  type ResultsFilter,
} from '@/mediacollector/title-block-container';
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
import {
  toFormImages,
  type MediaItemForm,
} from './collector-form/media-item-form-schema';
import { buildPNGExportImages } from './png-export-images';
import type { CollectorSearchProgress } from './collector-search-progress';
import {
  BuildingPngModal,
  CollectingCoversModal,
  SavingItemsModal,
} from './collector-loading-modal';
import DatabaseEditContext, {
  type DatabaseEditContextValue,
} from './database-edit-context';
import {
  countDatabaseItemChanges,
  diffDatabaseItem,
  type DatabaseItemEdit,
} from './database-item-changes';
import {
  buildDatabaseSaveFailureDisplayLines,
  markSuccessfulBlocksAsInDatabase,
  type DatabaseSaveFailureDisplayLine,
} from './database-save-error-display';

const sectionHeadingClassName =
  'm-0 text-[13px] font-extrabold tracking-[1px] text-[#6b645d] uppercase';
const actionButtonClassName =
  'flex h-12.5 cursor-pointer items-center justify-center gap-2.5 rounded-xl px-5.5 text-base font-extrabold disabled:cursor-not-allowed disabled:border-transparent disabled:bg-[#e3ddd5] disabled:text-[#5f5851]';
const emphasizedButtonClassName = `${actionButtonClassName} border border-transparent bg-[#2a4637] text-white`;
const quietButtonClassName = `${actionButtonClassName} border border-[#e0d9d0] bg-[#f7f4f0] text-[#6b645d]`;

type CollectorLoadingState =
  | {
      step: 'collecting';
      progress: CollectorSearchProgress;
      collectionList: CollectorFormData['collectionList'];
    }
  | { step: 'saving'; newItemCount: number }
  | {
      step: 'buildingPng';
      savedItemCount: number;
      customerName: string;
      orderNumber: string;
    };

interface HeaderActionProps {
  hasResults: boolean;
  lastSearchedTitles: string | null;
}

function useIsSearchNext({
  hasResults,
  lastSearchedTitles,
}: HeaderActionProps) {
  const collectionList = useWatch<CollectorFormData, 'collectionList'>({
    name: 'collectionList',
  });
  return {
    collectionList,
    isSearchNext:
      !hasResults || JSON.stringify(collectionList) !== lastSearchedTitles,
  };
}

function SearchTitlesButton({
  onClick,
  ...headerActionProps
}: HeaderActionProps & { onClick: () => void }) {
  const { collectionList, isSearchNext } = useIsSearchNext(headerActionProps);
  let titleCount = 0;
  for (const searchList of Object.values(collectionList ?? {})) {
    titleCount += searchList.length;
  }

  return (
    <button
      type="button"
      className={`${isSearchNext ? emphasizedButtonClassName : quietButtonClassName} mt-auto self-stretch`}
      onClick={onClick}
    >
      <SearchIcon fontSize="small" />
      {titleCount === 0
        ? 'Search'
        : `Search ${titleCount} ${titleCount === 1 ? 'title' : 'titles'}`}
    </button>
  );
}

function ExportButton(headerActionProps: HeaderActionProps) {
  const { isSearchNext } = useIsSearchNext(headerActionProps);
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
        className={
          isSearchNext ? quietButtonClassName : emphasizedButtonClassName
        }
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

function ResultsSummary({
  newCount,
  databaseCount,
}: {
  newCount: number;
  databaseCount: number;
}) {
  const collectedData = useWatch<CollectorFormData, 'collectedData'>({
    name: 'collectedData',
  });
  const { originalsByBlockId } = use(DatabaseEditContext);
  const blocksNeedingCover = (collectedData ?? []).filter(
    (block) => !block.isDatabase && !block.images.some((img) => img.selected),
  ).length;
  const editedDatabaseCount = (collectedData ?? []).filter((block) => {
    const original = block.isDatabase
      ? originalsByBlockId.get(block.blockID)
      : undefined;
    return (
      original !== undefined &&
      countDatabaseItemChanges(diffDatabaseItem(original, block)) > 0
    );
  }).length;

  return (
    <p className="m-0 font-semibold text-[#6b645d]">
      {newCount + databaseCount} results · {newCount} new · {databaseCount} from
      database
      {blocksNeedingCover > 0 ? (
        <>
          {' · '}
          <span className="font-extrabold text-[#9a3b26]">
            {blocksNeedingCover} still{' '}
            {blocksNeedingCover === 1 ? 'needs' : 'need'} a cover
          </span>
        </>
      ) : null}
      {editedDatabaseCount > 0 ? (
        <>
          {' · '}
          <span className="font-extrabold text-[#8a5a22]">
            {editedDatabaseCount} database{' '}
            {editedDatabaseCount === 1 ? 'item' : 'items'} edited
          </span>
        </>
      ) : null}
    </p>
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
  const [loadingState, setLoadingState] =
    useState<CollectorLoadingState | null>(null);
  const [informationalDialogText, setInformationalDialogText] = useState<
    string | null
  >(null);
  const [databaseSaveFailureLines, setDatabaseSaveFailureLines] = useState<
    DatabaseSaveFailureDisplayLine[]
  >([]);
  const [lastSearchedTitles, setLastSearchedTitles] = useState<string | null>(
    null,
  );
  const [resultsFilter, setResultsFilter] = useState<ResultsFilter>('all');
  const [databaseOriginals, setDatabaseOriginals] = useState<
    ReadonlyMap<string, MediaItemForm>
  >(() => new Map());
  const [unlockedBlockIds, setUnlockedBlockIds] = useState<ReadonlySet<string>>(
    () => new Set(),
  );
  const databaseEditContextValue = useMemo<DatabaseEditContextValue>(
    () => ({
      originalsByBlockId: databaseOriginals,
      unlockedBlockIds,
      setBlockUnlocked: (blockID, isUnlocked) => {
        setUnlockedBlockIds((current) => {
          const next = new Set(current);
          if (isUnlocked) {
            next.add(blockID);
          } else {
            next.delete(blockID);
          }
          return next;
        });
      },
    }),
    [databaseOriginals, unlockedBlockIds],
  );
  const databaseResultCount = fields.filter((field) => field.isDatabase).length;
  const newResultCount = fields.length - databaseResultCount;
  const resultsFilterOptions = [
    { filter: 'new', label: 'New', count: newResultCount },
    { filter: 'database', label: 'Database', count: databaseResultCount },
    { filter: 'all', label: 'All', count: fields.length },
  ] satisfies { filter: ResultsFilter; label: string; count: number }[];
  const visibleResultCount =
    resultsFilterOptions.find(({ filter }) => filter === resultsFilter)
      ?.count ?? 0;
  const [visibleMediaInputs, setVisibleMediaInputs] =
    useState<MediaVisibilityMap>({
      book: true,
      movie: false,
      videoGame: false,
      album: false,
    });

  const { mutateAsync: createPNG } = trpc.png.create.useMutation();
  const utils = trpc.useUtils();
  const { mutateAsync: collectMediaStream } =
    trpc.collect.collectMediaStream.useMutation();

  const applyCollectedBlocks = (
    blocks: MediaItemForm[],
    pngExportBlockIds: string[],
  ) => {
    const formBlocks = blocks.map((block) => ({
      ...block,
      images: toFormImages(block.images, block.blockInfo.spineColor),
    }));
    replace(formBlocks);
    setValue('pngExportBlockIds', pngExportBlockIds);
    setDatabaseOriginals(
      new Map(
        formBlocks
          .filter((block) => block.isDatabase)
          .map((block) => [block.blockID, block]),
      ),
    );
    setUnlockedBlockIds(new Set());
  };

  const handleCollectClick = async (): Promise<void> => {
    replace([]);
    setResultsFilter('all');
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

    setLoadingState({
      step: 'collecting',
      progress: { phase: 'collecting', searchedCount: 0, totalCount: count },
      collectionList,
    });

    try {
      const collectStream = await collectMediaStream(collectionList);
      for await (const event of collectStream) {
        if (event.type === 'progress') {
          setLoadingState({
            step: 'collecting',
            progress: event.progress,
            collectionList,
          });
          continue;
        }

        applyCollectedBlocks(event.blocks, event.pngExportBlockIds);
        setLastSearchedTitles(JSON.stringify(collectionList));
        const date = new Date().toLocaleDateString('en-CA', {
          timeZone: 'America/New_York',
        });
        void utils.database.getQueryCount.invalidate({ date });
      }
    } finally {
      setLoadingState(null);
    }
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
      return;
    }

    void handleSubmit(handleExportSubmit)();
  };

  const saveCollectedItems = async (
    formValues: CollectorFormData,
  ): Promise<boolean> => {
    const databaseEdits = formValues.collectedData.flatMap(
      (block): DatabaseItemEdit[] => {
        const original = block.isDatabase
          ? databaseOriginals.get(block.blockID)
          : undefined;
        return original &&
          countDatabaseItemChanges(diffDatabaseItem(original, block)) > 0
          ? [{ block, original }]
          : [];
      },
    );
    const result = await onSubmit(formValues, databaseEdits);

    const successfulBlockIds = new Set(
      result.filter((res) => res.success).map((res) => res.blockID),
    );
    const savedEdits = databaseEdits.filter(({ block }) =>
      successfulBlockIds.has(block.blockID),
    );
    if (savedEdits.length > 0) {
      setDatabaseOriginals((current) => {
        const next = new Map(current);
        for (const { block } of savedEdits) {
          next.set(block.blockID, block);
        }
        return next;
      });
      setUnlockedBlockIds((current) => {
        const next = new Set(current);
        for (const { block } of savedEdits) {
          next.delete(block.blockID);
        }
        return next;
      });
    }

    if (result.every((res) => res.success)) {
      return true;
    }

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
    return false;
  };

  const downloadPng = async (formValues: CollectorFormData): Promise<void> => {
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

  const handleExportSubmit = async (
    formValues: CollectorFormData,
  ): Promise<void> => {
    const newItemCount = formValues.collectedData.filter(
      (block) => !block.isDatabase,
    ).length;
    setLoadingState({ step: 'saving', newItemCount });

    try {
      const isSaved = await saveCollectedItems(formValues);
      if (!isSaved) {
        return;
      }

      setLoadingState({
        step: 'buildingPng',
        savedItemCount: newItemCount,
        customerName: formValues.customerName,
        orderNumber: formValues.orderNumber,
      });
      await downloadPng(formValues);
    } finally {
      setLoadingState(null);
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
    <DatabaseEditContext value={databaseEditContextValue}>
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
                <h2 className={sectionHeadingClassName}>
                  2 · Titles to search
                </h2>
                <MediaCheckboxes
                  visibility={visibleMediaInputs}
                  onToggle={handleToggleMediaInput}
                />
                <MediaInputs visibility={visibleMediaInputs} />
                <SearchTitlesButton
                  hasResults={fields.length > 0}
                  lastSearchedTitles={lastSearchedTitles}
                  onClick={() => {
                    void handleCollectClick();
                  }}
                />
              </div>

              <div className="flex flex-[1_1_250px] flex-col gap-3.5 border-l border-[#ece6de] pl-7">
                <h2 className={sectionHeadingClassName}>3 · Export</h2>
                <PNGFormatPicker />
                <ExportButton
                  hasResults={fields.length > 0}
                  lastSearchedTitles={lastSearchedTitles}
                />
                {submitErrorMessage ? (
                  <p className="m-0 text-xs font-bold text-red-600">
                    {submitErrorMessage}
                  </p>
                ) : null}
              </div>
            </div>
          </section>
        </form>
        {loadingState?.step === 'collecting' ? (
          <CollectingCoversModal
            progress={loadingState.progress}
            collectionList={loadingState.collectionList}
          />
        ) : null}
        {loadingState?.step === 'saving' ? (
          <SavingItemsModal newItemCount={loadingState.newItemCount} />
        ) : null}
        {loadingState?.step === 'buildingPng' ? (
          <BuildingPngModal
            savedItemCount={loadingState.savedItemCount}
            customerName={loadingState.customerName}
            orderNumber={loadingState.orderNumber}
          />
        ) : null}
        {databaseSaveFailureLines.length > 0 ? (
          <Dialog
            title="Error"
            onClose={() => setDatabaseSaveFailureLines([])}
          >
            <DatabaseSaveFailureBody failureLines={databaseSaveFailureLines} />
          </Dialog>
        ) : null}
        {informationalDialogText ? (
          <Dialog
            title="Error"
            onClose={() => setInformationalDialogText(null)}
            className="w-5/12"
          >
            <p>{informationalDialogText}</p>
          </Dialog>
        ) : null}
        {fields.length > 0 ? (
          <section className="flex flex-col gap-4.5">
            <div className="flex flex-wrap items-end justify-between gap-4">
              <div className="flex flex-col gap-1">
                <h2 className="m-0 text-[26px] leading-[1.1] font-extrabold">
                  Search results
                </h2>
                <ResultsSummary
                  newCount={newResultCount}
                  databaseCount={databaseResultCount}
                />
              </div>
              <div className="flex gap-1 rounded-full border border-[#e6e0d8] bg-white p-1">
                {resultsFilterOptions.map(({ filter, label, count }) => {
                  const isActive = resultsFilter === filter;
                  return (
                    <button
                      key={filter}
                      type="button"
                      aria-pressed={isActive}
                      onClick={() => setResultsFilter(filter)}
                      className={`h-9.5 cursor-pointer rounded-full border-0 px-4 font-bold ${
                        isActive
                          ? 'bg-[#2a4637] text-white'
                          : 'bg-transparent text-[#2b2826]'
                      }`}
                    >
                      {label}{' '}
                      <span
                        className={`font-extrabold ${
                          isActive
                            ? ''
                            : filter === 'new'
                              ? 'text-[#9a3b26]'
                              : 'text-[#6b645d]'
                        }`}
                      >
                        {count}
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>
            {visibleResultCount === 0 ? (
              <p className="m-0 font-semibold text-[#6b645d]">
                Nothing to show for this filter.
              </p>
            ) : (
              <TitleBlockContainer
                fields={fields}
                filter={resultsFilter}
                onDelete={remove}
                blockIdsWithErrors={[
                  ...new Set([...blockIdsWithErrors, ...schemaErrorBlockIds]),
                ]}
              />
            )}
          </section>
        ) : null}
      </main>
    </DatabaseEditContext>
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
