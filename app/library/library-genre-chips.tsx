import { use } from 'react';
import CloseIcon from '@mui/icons-material/Close';
import { useFormContext, useWatch } from 'react-hook-form';
import GenreContext from 'lib/context/genre-context';
import { Combobox } from '@/components/shared/combobox';
import type { MediaItemForm } from '@/mediacollector/collector-form/media-item-form-schema';

export default function LibraryGenreChips({
  isLocked,
  isChanged,
}: {
  isLocked: boolean;
  isChanged: boolean;
}) {
  const allGenres = use(GenreContext);
  const { control, getValues, setValue } = useFormContext<MediaItemForm>();
  const genres = useWatch({ control, name: 'blockInfo.genres' }) ?? [];

  const handleGenreClick = (genreText: string, checked: boolean) => {
    const currentGenres = getValues('blockInfo.genres') ?? [];
    setValue(
      'blockInfo.genres',
      checked
        ? [...currentGenres, genreText]
        : currentGenres.filter((genre) => genre !== genreText),
    );
  };

  return (
    <div
      className={`flex flex-wrap items-center gap-1.5 rounded-lg ${
        isChanged
          ? 'outline-[1.5px] outline-offset-[3px] outline-[#c9973f]'
          : ''
      }`}
    >
      {genres.map((genre) => (
        <span
          key={genre}
          className={`inline-flex h-7 items-center gap-1 rounded-full bg-[#efebe5] text-xs font-bold ${
            isLocked ? 'px-2.5' : 'pr-1 pl-2.5'
          }`}
        >
          {genre}
          {isLocked ? null : (
            <button
              type="button"
              onClick={() => handleGenreClick(genre, false)}
              className="flex h-5.5 w-5.5 cursor-pointer items-center justify-center rounded-full border-0 bg-transparent p-0 text-[#6b645d]"
            >
              <CloseIcon sx={{ fontSize: 14 }} />
            </button>
          )}
        </span>
      ))}
      {isLocked ? null : (
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
      )}
    </div>
  );
}
