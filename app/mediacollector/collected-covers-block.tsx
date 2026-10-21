import { memo, useContext } from 'react';
import IconButton from '@mui/material/IconButton';
import DeleteIcon from '@mui/icons-material/Delete';
import CBBImages from '@/mediacollector/cbb-images';
import CollectedItemTextField from '@/mediacollector/collected-item-text-field';
import GenreContext from 'lib/context/genre-context';
import GenreCheckboxes from '@/mediacollector/genre-checkboxes';
import { useFormContext } from 'react-hook-form';
import { CollectorFormData } from './collector-form/collector-form-schema';
import { blockClasses, icons } from 'lib/constants/type-block-styles';
import { MediaType } from 'lib/constants/media-types';

export interface CollectedCoversBlockProps {
  index: number;
  type: MediaType;
  isDatabase: boolean;
  blockID: string;
  spineColor: string;
  onDelete: (index: number) => void;
  hasError: boolean;
}

const CollectedCoversBlock = memo(function CollectedCoversBlock({
  hasError,
  index,
  type,
  isDatabase,
  blockID,
  spineColor,
  onDelete,
}: CollectedCoversBlockProps) {
  const allGenres = useContext(GenreContext);
  const { getValues, setValue, watch } = useFormContext<CollectorFormData>();
  const genres = watch(`collectedData.${index}.blockInfo.genres`) ?? [];

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
    <div
      className={`relative flex h-full w-full flex-col items-center gap-1.5 rounded-lg shadow-[5px_5px_30px_rgba(0,0,0,0.3)] ${hasError ? blockClasses.hasError : blockClasses[type]}`}
    >
      {isDatabase && (
        <p className="absolute top-9 left-1 m-0 rounded-sm border-2 border-black bg-gray-700 p-1.25 tracking-wider text-white">
          Database
        </p>
      )}
      <div className="absolute top-1 left-1">{icons[type]}</div>
      <div className="absolute top-1 right-1">block: {index + 1}</div>

      <IconButton
        aria-label="delete"
        sx={{
          position: 'absolute',
          bottom: '4px',
          right: '4px',
          padding: '0',
          color: type === 'movie' ? 'white' : '',
        }}
        onClick={() => onDelete(index)}
      >
        <DeleteIcon />
      </IconButton>

      <CBBImages
        index={index}
        blockID={blockID}
        type={type}
        isDatabase={isDatabase}
        spineColor={spineColor}
      />

      <CollectedItemTextField
        name="title"
        label="Title"
        type={type}
        index={index}
      />
      {type === 'book' ? (
        <>
          <CollectedItemTextField
            name="author"
            label="Author"
            type={type}
            index={index}
          />
          <CollectedItemTextField
            name="pubYear"
            label="Pub Year"
            type={type}
            index={index}
          />
          <CollectedItemTextField
            name="pageCount"
            label="Page Count"
            type={type}
            index={index}
          />
        </>
      ) : null}
      {type === 'book' ? (
        <GenreCheckboxes
          allGenres={allGenres}
          handleGenreClick={handleGenreClick}
          blockGenres={genres}
        />
      ) : null}
    </div>
  );
});

export default CollectedCoversBlock;
