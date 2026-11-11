'use client';
import { memo, use, useEffect, useRef, useState, useTransition } from 'react';
import DeleteIcon from '@mui/icons-material/DeleteOutline';
import LockIcon from '@mui/icons-material/LockOutlined';
import LockOpenIcon from '@mui/icons-material/LockOpenOutlined';
import ReplayIcon from '@mui/icons-material/Replay';
import { trpc } from 'lib/trpc/client';
import { useLibraryPageContext } from 'lib/context/library-page-context';
import LibraryPendingSavesContext from 'lib/context/library-pending-saves-context';
import { PostSavedMediaItem } from 'lib/interfaces/global-interfaces';
import { MEDIA_TYPE_DEFINITIONS, MEDIA_TYPES } from 'lib/constants/media-types';
import { Form } from '@/components/ui/form';
import {
  convertMediaItemToForm,
  type MediaItemForm,
} from '@/mediacollector/collector-form/media-item-form-schema';
import {
  movieFormatLabel,
  typeChipClassNames,
} from '@/mediacollector/collected-covers-block';
import type { CollectedFieldEditState } from '@/mediacollector/collected-item-text-field';
import {
  CoverImage,
  collectorCoverAspectClassName,
} from '@/mediacollector/cover-tile';
import {
  countDatabaseItemChanges,
  diffDatabaseItem,
  restoreDatabaseItem,
} from '@/mediacollector/database-item-changes';
import DeleteConfirmDialog from './delete-confirm-dialog';
import { LibraryGenreRow } from './library-genres';
import LibraryItemCovers from './library-item-covers';
import LibraryItemTextField from './library-item-text-field';
import LibrarySpineColorRow from './library-spine-color-row';
import { useLibraryItemForm } from './use-library-item-form';

const pillClassName = 'rounded-full px-2.5 py-1 text-xs font-extrabold';
const footerButtonClassName =
  'flex h-9 cursor-pointer items-center gap-1.5 rounded-lg border px-3 text-[13px] font-bold disabled:cursor-not-allowed';

function LibraryCard({ item }: { item: MediaItemForm }) {
  const { refreshLibrary } = useLibraryPageContext();
  const [isLocked, setIsLocked] = useState(true);
  const [baseline, setBaseline] = useState(item);
  const [isConfirmingDelete, setIsConfirmingDelete] = useState(false);
  const [deleteError, setDeleteError] = useState<string | null>(null);
  const [isDeleting, startDeleting] = useTransition();
  const { form, formId, onSubmit, submitError } = useLibraryItemForm({
    item,
    onClose: () => setIsLocked(true),
    onEditSuccess: (saved) => {
      setBaseline(saved);
      setIsLocked(true);
    },
  });
  const { mutateAsync: databaseDelete } = trpc.database.delete.useMutation();

  const current = form.watch();
  const { type, images, blockInfo } = current;
  const changes = diffDatabaseItem(baseline, current);
  const changeCount = countDatabaseItemChanges(changes);
  const isSaving = form.formState.isSubmitting;
  const hasPendingChanges = changeCount > 0;

  const registerPendingSave = use(LibraryPendingSavesContext);
  const saveRef = useRef(() => Promise.resolve());
  useEffect(() => {
    saveRef.current = form.handleSubmit(onSubmit);
  });
  useEffect(() => {
    if (!hasPendingChanges) {
      return;
    }
    return registerPendingSave(item.blockID, () => saveRef.current());
  }, [hasPendingChanges, item.blockID, registerPendingSave]);

  const fieldEditState = (
    isChanged: boolean,
  ): CollectedFieldEditState | undefined =>
    isLocked ? 'readOnly' : isChanged ? 'changed' : undefined;

  const selectedImage = images.find((image) => image.selected);
  const originalDefaultUrl = baseline.images.find(
    (image) => image.isDefault,
  )?.url;
  const typeLabel =
    MEDIA_TYPE_DEFINITIONS.find(({ mediaType }) => mediaType === type)?.label ??
    type;
  const mainCoverSizeClassName = type === 'album' ? 'h-30 w-30' : 'h-45 w-30';

  const handleDelete = () =>
    startDeleting(async () => {
      setDeleteError(null);
      try {
        await databaseDelete({ id: item.blockID, type });
        await refreshLibrary();
      } catch {
        setDeleteError('Couldn’t delete. Try again.');
        return;
      }
      setIsConfirmingDelete(false);
    });

  return (
    <Form {...form}>
      <form id={formId} onSubmit={form.handleSubmit(onSubmit)}>
        <article
          className={`relative flex flex-col overflow-hidden rounded-2xl border bg-white shadow-[0_4px_18px_rgba(60,50,40,0.06)] ${
            changeCount > 0
              ? 'border-[#c9973f] ring-1 ring-[#c9973f]'
              : 'border-[#e6e0d8]'
          }`}
        >
          <div className="flex items-center justify-between gap-2 px-4.5 pt-3.5">
            <div className="flex items-center gap-1.5">
              <span className={`${pillClassName} ${typeChipClassNames[type]}`}>
                {typeLabel}
              </span>
              {type === 'movie' ? (
                <span className={`${pillClassName} ${typeChipClassNames.book}`}>
                  {movieFormatLabel(blockInfo.movieFormat)}
                </span>
              ) : null}
            </div>
            <span
              role="img"
              aria-label={
                isLocked
                  ? 'Locked'
                  : changeCount > 0
                    ? `Unlocked, ${changeCount} unsaved ${changeCount === 1 ? 'change' : 'changes'}`
                    : 'Unlocked'
              }
              className={`inline-flex h-6.5 items-center gap-1 rounded-full px-2 ${
                changeCount > 0
                  ? 'bg-[#f6ead6] text-[#6b4a14]'
                  : 'bg-[#e3e9df] text-[#3f5637]'
              }`}
            >
              {isLocked ? (
                <LockIcon sx={{ fontSize: 15 }} />
              ) : (
                <LockOpenIcon sx={{ fontSize: 15 }} />
              )}
              {changeCount > 0 ? (
                <span className="flex h-4 min-w-4 items-center justify-center rounded-full bg-[#6b4a14] px-1 text-[10px] font-extrabold text-white">
                  {changeCount}
                </span>
              ) : null}
            </span>
          </div>

          {type === 'movie' ? (
            <div className="mb-3.5 flex flex-col gap-3 px-4.5 pt-3.5">
              <LibraryItemTextField
                name="title"
                label="Title"
                editState={fieldEditState(changes.title)}
              />
              {selectedImage ? (
                <div
                  className={`relative w-full overflow-hidden rounded-md ${collectorCoverAspectClassName('movie')}`}
                >
                  <CoverImage
                    url={selectedImage.url}
                    alt={`${blockInfo.title} cover`}
                  />
                </div>
              ) : null}
            </div>
          ) : (
            <div
              className={`flex gap-4.5 px-4.5 py-3.5 ${
                type === 'book' ? 'items-center' : 'items-start'
              }`}
            >
              {selectedImage ? (
                <div className="flex shrink-0 overflow-hidden rounded-r-md">
                  {type !== 'album' ? (
                    <div
                      className="w-2.5 shrink-0"
                      style={{ backgroundColor: blockInfo.spineColor }}
                    />
                  ) : null}
                  <div className={`relative ${mainCoverSizeClassName}`}>
                    <CoverImage
                      url={selectedImage.url}
                      alt={`${blockInfo.title} cover`}
                    />
                  </div>
                </div>
              ) : null}

              <div className="flex min-w-0 flex-1 flex-col gap-2.5">
                <LibraryItemTextField
                  name="title"
                  label="Title"
                  editState={fieldEditState(changes.title)}
                />
                {type === 'book' ? (
                  <>
                    <LibraryItemTextField
                      name="author"
                      label="Author"
                      editState={fieldEditState(changes.author)}
                    />
                    <div className="grid grid-cols-2 gap-2.5">
                      <LibraryItemTextField
                        name="pubYear"
                        label="Published"
                        editState={fieldEditState(changes.pubYear)}
                      />
                      <LibraryItemTextField
                        name="pageCount"
                        label="Pages"
                        editState={fieldEditState(changes.pageCount)}
                      />
                    </div>
                  </>
                ) : null}
              </div>
            </div>
          )}

          {type === 'book' ? (
            <LibraryGenreRow
              isEditable={!isLocked}
              isChanged={changes.genres}
            />
          ) : null}

          <LibraryItemCovers
            type={type}
            blockID={item.blockID}
            title={blockInfo.title}
            isLocked={isLocked}
            originalDefaultUrl={originalDefaultUrl}
            isCoverListChanged={changes.coverList}
          />

          {type !== 'album' && type !== 'movie' ? (
            <LibrarySpineColorRow
              isLocked={isLocked}
              isChanged={
                selectedImage !== undefined &&
                changes.changedSpineColorUrls.includes(selectedImage.url)
              }
            />
          ) : null}

          {submitError ? (
            <p className="mx-4.5 mt-0 mb-3 text-xs font-bold text-red-600">
              {submitError}
            </p>
          ) : null}

          <div className="mt-auto flex items-center justify-between gap-2 border-t border-[#efe9e2] px-2.5 py-2">
            <div className="flex items-center gap-2">
              {isLocked ? (
                <button
                  type="button"
                  onClick={() => setIsLocked(false)}
                  className={`${footerButtonClassName} border-[#e0d9d0] bg-white text-[#2b2826]`}
                >
                  <LockOpenIcon sx={{ fontSize: 16 }} />
                  Unlock
                </button>
              ) : changeCount > 0 ? (
                <button
                  type="button"
                  disabled={isSaving}
                  onClick={() =>
                    form.reset(restoreDatabaseItem(baseline, form.getValues()))
                  }
                  className={`${footerButtonClassName} border-[#c9973f] bg-[#fffaf0] text-[#6b4a14]`}
                >
                  <ReplayIcon sx={{ fontSize: 16 }} />
                  Clear changes
                </button>
              ) : (
                <button
                  type="button"
                  onClick={() => setIsLocked(true)}
                  className={`${footerButtonClassName} border-[#e0d9d0] bg-white text-[#2b2826]`}
                >
                  Cancel
                </button>
              )}
              {isLocked ? null : (
                <button
                  type="submit"
                  disabled={changeCount === 0 || isSaving}
                  className={`${footerButtonClassName} border-transparent bg-[#2a4637] text-white disabled:bg-[#e3ddd5] disabled:text-[#5f5851]`}
                >
                  {isSaving ? 'Saving…' : 'Save changes'}
                </button>
              )}
            </div>
            <button
              type="button"
              onClick={() => {
                setDeleteError(null);
                setIsConfirmingDelete(true);
              }}
              className="flex h-9 cursor-pointer items-center gap-1.5 rounded-lg border-0 bg-transparent px-2.5 text-[13px] font-bold text-[#9a3b26] hover:bg-[#f7f4f0]"
            >
              <DeleteIcon fontSize="small" />
              Delete
            </button>
          </div>
        </article>
      </form>
      {isConfirmingDelete ? (
        <DeleteConfirmDialog
          item={baseline}
          isDeleting={isDeleting}
          error={deleteError}
          onCancel={() => setIsConfirmingDelete(false)}
          onConfirm={handleDelete}
        />
      ) : null}
    </Form>
  );
}

export interface LibraryItemProps {
  item: PostSavedMediaItem;
}

const LibraryItem = memo(function LibraryItem({ item }: LibraryItemProps) {
  const {
    libraryList: { type: displayedListType },
  } = useLibraryPageContext();
  const itemType =
    MEDIA_TYPES.find((mediaType) => mediaType === item.mediaType) ??
    displayedListType;

  return (
    <LibraryCard
      item={convertMediaItemToForm({
        item,
        type: itemType,
      })}
    />
  );
});

export default LibraryItem;
