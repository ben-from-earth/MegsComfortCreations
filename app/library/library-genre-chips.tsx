import { useFormContext, useWatch } from 'react-hook-form';
import { GenreChipList, GenreRow } from '@/components/shared/genre-row';
import type { MediaItemForm } from '@/mediacollector/collector-form/media-item-form-schema';

function useFormGenres() {
  const { control, getValues, setValue } = useFormContext<MediaItemForm>();
  const genres = useWatch({ control, name: 'blockInfo.genres' }) ?? [];
  const currentGenres = () => getValues('blockInfo.genres') ?? [];

  return {
    genres,
    onAdd: (genreText: string) =>
      setValue('blockInfo.genres', [...currentGenres(), genreText]),
    onRemove: (genreText: string) =>
      setValue(
        'blockInfo.genres',
        currentGenres().filter((genre) => genre !== genreText),
      ),
  };
}

export function LibraryGenreRow({
  isEditable,
  isChanged,
}: {
  isEditable: boolean;
  isChanged: boolean;
}) {
  return (
    <GenreRow
      {...useFormGenres()}
      isEditable={isEditable}
      isChanged={isChanged}
    />
  );
}

export function AddBookGenreChips() {
  return <GenreChipList {...useFormGenres()} />;
}
