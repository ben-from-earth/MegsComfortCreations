import { memo, use } from 'react';
import DeleteIcon from '@mui/icons-material/DeleteOutline';
import LockIcon from '@mui/icons-material/LockOutlined';
import LockOpenIcon from '@mui/icons-material/LockOpenOutlined';
import ReplayIcon from '@mui/icons-material/Replay';
import { useFormContext, useWatch } from 'react-hook-form';
import { GenreRow } from '@/components/shared/genre-row';
import CBBImages from '@/mediacollector/cbb-images';
import CollectedItemTextField, {
  type CollectedFieldEditState,
} from '@/mediacollector/collected-item-text-field';
import { CollectorFormData } from './collector-form/collector-form-schema';
import { MEDIA_TYPE_DEFINITIONS, MediaType } from 'lib/constants/media-types';
import SpineColorRow from './spine-color-row';
import { CoverImage, collectorCoverAspectClassName } from './cover-tile';
import DatabaseEditContext from './database-edit-context';
import {
  countDatabaseItemChanges,
  diffDatabaseItem,
  restoreDatabaseItem,
} from './database-item-changes';

const pillClassName = 'rounded-full px-2.5 py-1 text-xs font-extrabold';
const footerButtonClassName =
  'flex h-9 cursor-pointer items-center gap-1.5 rounded-lg border px-3 text-[13px] font-bold';

export const typeChipClassNames: Record<MediaType, string> = {
  book: 'bg-[#efebe5] text-[#4a443f]',
  movie: 'bg-[#3d3935] text-white',
  album: 'bg-[#e4ebea] text-[#2f5452]',
  videoGame: 'bg-[#e3e9df] text-[#3f5637]',
};

export interface CollectedCoversBlockProps {
  index: number;
  type: MediaType;
  isDatabase: boolean;
  blockID: string;
  onDelete: (index: number) => void;
  hasError: boolean;
}

const CollectedCoversBlock = memo(function CollectedCoversBlock({
  hasError,
  index,
  type,
  isDatabase,
  blockID,
  onDelete,
}: CollectedCoversBlockProps) {
  const { originalsByBlockId, unlockedBlockIds, setBlockUnlocked } =
    use(DatabaseEditContext);
  const { control, getValues, setValue } = useFormContext<CollectorFormData>();
  const block = useWatch({ control, name: `collectedData.${index}` });
  const images = block?.images ?? [];
  const title = block?.blockInfo.title ?? '';
  const spineColor = block?.blockInfo.spineColor;
  const genres = block?.blockInfo.genres ?? [];

  const original = isDatabase ? originalsByBlockId.get(blockID) : undefined;
  const changes = original && block ? diffDatabaseItem(original, block) : null;
  const changeCount = changes ? countDatabaseItemChanges(changes) : 0;
  const isLocked = isDatabase && !unlockedBlockIds.has(blockID);
  const fieldEditState = (
    isChanged: boolean | undefined,
  ): CollectedFieldEditState | undefined =>
    isLocked ? 'readOnly' : isChanged ? 'changed' : undefined;

  const originalDefaultUrl = original?.images.find(
    (image) => image.isDefault,
  )?.url;
  const selectedImage = images.find((image) => image.selected);
  const needsCover = !isDatabase && !selectedImage;
  const typeLabel =
    MEDIA_TYPE_DEFINITIONS.find(({ mediaType }) => mediaType === type)?.label ??
    type;
  const mainCoverSizeClassName = type === 'album' ? 'h-30 w-30' : 'h-45 w-30';

  const newItemBadge = needsCover
    ? { text: 'New · pick a cover', className: 'bg-[#f7e3dc] text-[#9a3b26]' }
    : { text: 'New · ready', className: 'bg-[#f0e7d8] text-[#6b5532]' };
  const handleGenreClick = (genreText: string, checked: boolean) => {
    const currentGenres =
      getValues(`collectedData.${index}.blockInfo.genres`) ?? [];
    setValue(
      `collectedData.${index}.blockInfo.genres`,
      checked
        ? [...currentGenres, genreText]
        : currentGenres.filter((genre) => genre !== genreText),
    );
  };
  const handleClearChanges = () => {
    if (!original) {
      return;
    }
    setValue(
      `collectedData.${index}`,
      restoreDatabaseItem(original, getValues(`collectedData.${index}`)),
    );
    setBlockUnlocked(blockID, false);
  };

  return (
    <article
      className={`relative flex flex-col overflow-hidden rounded-2xl border bg-white shadow-[0_4px_18px_rgba(60,50,40,0.06)] ${
        hasError || needsCover
          ? 'border-[#b0533a] ring-1 ring-[#b0533a]'
          : changeCount > 0
            ? 'border-[#c9973f] ring-1 ring-[#c9973f]'
            : 'border-[#e6e0d8]'
      }`}
    >
      <div className="flex items-center justify-between gap-2 px-4.5 pt-3.5">
        <span className={`${pillClassName} ${typeChipClassNames[type]}`}>
          {typeLabel} · #{index + 1}
        </span>
        {isDatabase ? (
          <span
            className={`${pillClassName} inline-flex items-center gap-1 ${
              changeCount > 0
                ? 'bg-[#f6ead6] text-[#6b4a14]'
                : 'bg-[#e3e9df] text-[#3f5637]'
            }`}
          >
            {isLocked ? (
              <LockIcon sx={{ fontSize: 13 }} />
            ) : (
              <LockOpenIcon sx={{ fontSize: 13 }} />
            )}
            Database
            {changeCount > 0 ? (
              <span className="flex h-4 min-w-4 items-center justify-center rounded-full bg-[#6b4a14] px-1 text-[10px] text-white">
                {changeCount}
              </span>
            ) : null}
          </span>
        ) : (
          <span className={`${pillClassName} ${newItemBadge.className}`}>
            {newItemBadge.text}
          </span>
        )}
      </div>

      {type === 'movie' ? (
        <div className="mb-3.5 flex flex-col gap-3 px-4.5 pt-3.5">
          <CollectedItemTextField
            name="title"
            label="Title"
            index={index}
            editState={fieldEditState(changes?.title)}
          />
          {selectedImage ? (
            <div
              className={`relative w-full overflow-hidden rounded-md ${collectorCoverAspectClassName('movie')}`}
            >
              <CoverImage url={selectedImage.url} alt={`${title} cover`} />
            </div>
          ) : (
            <div
              className={`flex w-full items-center justify-center rounded-md border-2 border-dashed border-[#b0533a] bg-[#f9eee9] p-2 text-center text-[13px] font-extrabold text-[#9a3b26] ${collectorCoverAspectClassName('movie')}`}
            >
              Choose a cover below
            </div>
          )}
        </div>
      ) : (
        <>
          <div className="flex items-center gap-4.5 px-4.5 py-3.5">
            {selectedImage ? (
              <div className="flex shrink-0 overflow-hidden rounded-r-md">
                {type !== 'album' ? (
                  <div
                    className="w-2.5 shrink-0"
                    style={{ backgroundColor: spineColor }}
                  />
                ) : null}
                <div className={`relative ${mainCoverSizeClassName}`}>
                  <CoverImage url={selectedImage.url} alt={`${title} cover`} />
                </div>
              </div>
            ) : (
              <div
                className={`flex shrink-0 items-center justify-center rounded-md border-2 border-dashed border-[#b0533a] bg-[#f9eee9] p-2 text-center text-[13px] font-extrabold text-[#9a3b26] ${mainCoverSizeClassName}`}
              >
                Choose a cover below
              </div>
            )}

            <div className="flex min-w-0 flex-1 flex-col gap-2.5">
              <CollectedItemTextField
                name="title"
                label="Title"
                index={index}
                editState={fieldEditState(changes?.title)}
              />
              {type === 'book' ? (
                <>
                  <CollectedItemTextField
                    name="author"
                    label="Author"
                    index={index}
                    editState={fieldEditState(changes?.author)}
                  />
                  <div className="grid grid-cols-2 gap-2.5">
                    <CollectedItemTextField
                      name="pubYear"
                      label="Published"
                      index={index}
                      editState={fieldEditState(changes?.pubYear)}
                    />
                    <CollectedItemTextField
                      name="pageCount"
                      label="Pages"
                      index={index}
                      editState={fieldEditState(changes?.pageCount)}
                    />
                  </div>
                </>
              ) : null}
            </div>
          </div>

          {type === 'book' ? (
            <GenreRow
              genres={genres}
              isEditable={!isLocked}
              isChanged={changes?.genres ?? false}
              onAdd={(genre) => handleGenreClick(genre, true)}
              onRemove={(genre) => handleGenreClick(genre, false)}
            />
          ) : (
            <div className="mx-4.5 mb-3.5 h-10" />
          )}
        </>
      )}

      <CBBImages
        index={index}
        blockID={blockID}
        type={type}
        isDatabase={isDatabase}
        isLocked={isLocked}
        originalDefaultUrl={originalDefaultUrl}
        title={title}
      />

      {type !== 'album' && type !== 'movie' ? (
        <SpineColorRow
          index={index}
          isLocked={isLocked}
          isChanged={
            selectedImage !== undefined &&
            (changes?.changedSpineColorUrls.includes(selectedImage.url) ??
              false)
          }
        />
      ) : null}

      <div className="mt-auto flex items-center justify-between gap-2 border-t border-[#efe9e2] px-2.5 py-2">
        {isDatabase ? (
          isLocked ? (
            <button
              type="button"
              onClick={() => setBlockUnlocked(blockID, true)}
              className={`${footerButtonClassName} border-[#e0d9d0] bg-white text-[#2b2826]`}
            >
              <LockOpenIcon sx={{ fontSize: 16 }} />
              Unlock editing
            </button>
          ) : changeCount > 0 ? (
            <button
              type="button"
              onClick={handleClearChanges}
              className={`${footerButtonClassName} border-[#c9973f] bg-[#fffaf0] text-[#6b4a14]`}
            >
              <ReplayIcon sx={{ fontSize: 16 }} />
              Clear changes
            </button>
          ) : (
            <button
              type="button"
              onClick={() => setBlockUnlocked(blockID, false)}
              className={`${footerButtonClassName} border-[#e0d9d0] bg-white text-[#2b2826]`}
            >
              Cancel
            </button>
          )
        ) : (
          <span />
        )}
        <button
          type="button"
          onClick={() => onDelete(index)}
          className="flex h-9 cursor-pointer items-center gap-1.5 rounded-lg border-0 bg-transparent px-2.5 text-[13px] font-bold text-[#6b645d] hover:bg-[#f7f4f0]"
        >
          <DeleteIcon fontSize="small" />
          Remove
        </button>
      </div>
    </article>
  );
});

export default CollectedCoversBlock;
