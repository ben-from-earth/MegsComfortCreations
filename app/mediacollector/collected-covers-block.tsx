import { memo, use } from 'react';
import CloseIcon from '@mui/icons-material/Close';
import DeleteIcon from '@mui/icons-material/DeleteOutline';
import { useFormContext, useWatch } from 'react-hook-form';
import GenreContext from 'lib/context/genre-context';
import { Combobox } from '@/components/shared/combobox';
import CBBImages from '@/mediacollector/cbb-images';
import CollectedItemTextField from '@/mediacollector/collected-item-text-field';
import { CollectorFormData } from './collector-form/collector-form-schema';
import { MEDIA_TYPE_DEFINITIONS, MediaType } from 'lib/constants/media-types';
import SpineColorRow from './spine-color-row';
import { CoverImage } from './cover-tile';

const pillClassName = 'rounded-full px-2.5 py-1 text-xs font-extrabold';

const typeChipClassNames: Record<MediaType, string> = {
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
  const allGenres = use(GenreContext);
  const { control, getValues, setValue } = useFormContext<CollectorFormData>();
  const images =
    useWatch({ control, name: `collectedData.${index}.images` }) ?? [];
  const title =
    useWatch({ control, name: `collectedData.${index}.blockInfo.title` }) ??
    '';
  const spineColor = useWatch({
    control,
    name: `collectedData.${index}.blockInfo.spineColor`,
  });
  const genres =
    useWatch({ control, name: `collectedData.${index}.blockInfo.genres` }) ??
    [];

  const selectedImage = images.find((image) => image.selected);
  const needsCover = !isDatabase && !selectedImage;
  const typeLabel =
    MEDIA_TYPE_DEFINITIONS.find(({ mediaType }) => mediaType === type)
      ?.label ?? type;
  const mainCoverSizeClassName = type === 'album' ? 'h-30 w-30' : 'h-45 w-30';

  const badge = isDatabase
    ? { text: 'Database', className: 'bg-[#e3e9df] text-[#3f5637]' }
    : needsCover
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

  return (
    <article
      className={`relative flex flex-col overflow-hidden rounded-2xl bg-white shadow-[0_4px_18px_rgba(60,50,40,0.06)] ${
        hasError || needsCover
          ? 'border-2 border-[#b0533a]'
          : 'border border-[#e6e0d8]'
      }`}
    >
      <div className="flex items-center justify-between gap-2 px-4.5 pt-3.5">
        <span className={`${pillClassName} ${typeChipClassNames[type]}`}>
          {typeLabel} · #{index + 1}
        </span>
        <span className={`${pillClassName} ${badge.className}`}>
          {badge.text}
        </span>
      </div>

      <div className="flex items-center gap-4.5 px-4.5 py-3.5">
        {selectedImage ? (
          <div
            className={`relative shrink-0 overflow-hidden rounded-md shadow-[0_0_0_3px_#ffffff,0_0_0_5px_#2a4637] ${mainCoverSizeClassName}`}
          >
            <CoverImage url={selectedImage.url} alt={`${title} cover`} />
            {type !== 'album' ? (
              <div
                className="absolute inset-y-0 left-0 w-2.5"
                style={{ backgroundColor: spineColor }}
              />
            ) : null}
            {isDatabase && selectedImage.isDefault ? (
              <span className="absolute top-2 right-2 rounded-full bg-white px-2 py-0.75 text-[10px] font-extrabold tracking-[0.5px] text-[#2b2826] uppercase">
                Default
              </span>
            ) : null}
          </div>
        ) : (
          <div
            className={`flex shrink-0 items-center justify-center rounded-md border-2 border-dashed border-[#b0533a] bg-[#f9eee9] p-2 text-center text-[13px] font-extrabold text-[#9a3b26] ${mainCoverSizeClassName}`}
          >
            Choose a cover below
          </div>
        )}

        <div className="flex min-w-0 flex-1 flex-col gap-2.5">
          <CollectedItemTextField name="title" label="Title" index={index} />
          {type === 'book' ? (
            <>
              <CollectedItemTextField
                name="author"
                label="Author"
                index={index}
              />
              <div className="grid grid-cols-2 gap-2.5">
                <CollectedItemTextField
                  name="pubYear"
                  label="Published"
                  index={index}
                />
                <CollectedItemTextField
                  name="pageCount"
                  label="Pages"
                  index={index}
                />
              </div>
              <div className="flex flex-wrap items-center gap-1.5">
                {genres.map((genre) => (
                  <span
                    key={genre}
                    className="inline-flex h-7 items-center gap-1 rounded-full bg-[#efebe5] pr-1 pl-2.5 text-xs font-bold"
                  >
                    {genre}
                    <button
                      type="button"
                      onClick={() => handleGenreClick(genre, false)}
                      className="flex h-5.5 w-5.5 cursor-pointer items-center justify-center rounded-full border-0 bg-transparent p-0 text-[#6b645d]"
                    >
                      <CloseIcon sx={{ fontSize: 14 }} />
                    </button>
                  </span>
                ))}
                <Combobox
                  label="genres"
                  items={allGenres
                    .filter((genre) => !genres.includes(genre))
                    .map((genre) => ({ value: genre }))}
                  onSelect={(genre) => handleGenreClick(genre, true)}
                  trigger={
                    <button
                      type="button"
                      className="h-7 cursor-pointer rounded-full border border-dashed border-[#b3aaa0] bg-transparent px-2.5 text-xs font-bold text-[#6b645d]"
                    >
                      + Genre
                    </button>
                  }
                />
              </div>
            </>
          ) : null}
        </div>
      </div>

      <CBBImages
        index={index}
        blockID={blockID}
        type={type}
        isDatabase={isDatabase}
        title={title}
      />

      {type !== 'album' ? (
        <SpineColorRow index={index} />
      ) : null}

      <div className="mt-auto flex items-center justify-end border-t border-[#efe9e2] px-2.5 py-2">
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
