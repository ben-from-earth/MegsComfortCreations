// react
import { memo, useContext } from 'react';

//import icons and items from Material UI
import BookIcon from '@mui/icons-material/BookTwoTone';
import MovieIcon from '@mui/icons-material/LocalMoviesTwoTone';
import VideoGameIcon from '@mui/icons-material/VideogameAssetTwoTone';
import AlbumIcon from '@mui/icons-material/AlbumTwoTone';
import IconButton from '@mui/material/IconButton';
import DeleteIcon from '@mui/icons-material/Delete';

// components
import CBBImages from '@/mediacollector/CBBImages';
import MyTextArea from '@/shared/MyTextArea';
import GenreContext from 'lib/context/GenreContext';
import GenreCheckboxes from '@/mediacollector/GenreCheckboxes';
import { useFormContext } from 'react-hook-form';

import {
  CollectedBlockInformation,
  CollectorFormData,
} from './collector-form/collectorFormSchema';

export interface CollectedCoversBlockProps {
  index: number;
  info: CollectedBlockInformation;
  handleDeleteBlock: (blockID: string) => void;
  hasError: boolean;
}

declare global {
  interface Window {
    EyeDropper?: {
      new (): {
        open: () => Promise<{ sRGBHex: string }>;
      };
    };
  }
}

//styling of the block itself based on type
export const blockClasses = {
  book: 'bg-[#e1b3b5]',
  movie: 'bg-[#323b43] text-white',
  album: 'bg-[#7fa5a3]',
  videoGame: 'bg-[#98ab88]',
  hasError: 'bg-[#E86C54]',
};

const CollectedCoversBlock = memo(function CollectedCoversBlock({
  hasError,
  index,
  info,
  handleDeleteBlock,
}: CollectedCoversBlockProps) {
  const {
    type,
    blockInfo: { title, genres = [] },
    blockID,
    isDatabase,
  } = info;

  //get genres for checkbox population
  const allGenres = useContext(GenreContext);

  const { watch, setValue } = useFormContext<CollectorFormData>();
  const collectedData = watch('collectedData');
  const block = collectedData.find((block) => block.blockID === blockID);
  if (!block) {
    return null;
  }

  //setup connection to redux slice

  //establish variables for icons
  const icons = {
    book: <BookIcon />,
    movie: <MovieIcon />,
    videoGame: <VideoGameIcon />,
    album: <AlbumIcon />,
  };

  //div to pick a color for the spine.
  //this is used in png creation and is required for each media type in the database

  //if genre is clicked we add it to the data associated with the block and remove if unchecked
  const handleGenreClick = (genreText: string, checked: boolean) => {
    if (checked) {
      const newGenres = [...genres, genreText];
      const newBlock = {
        ...info,
        blockInfo: { ...info.blockInfo, genres: newGenres },
      };
      setValue(`collectedData.${index}`, newBlock);
    } else {
      const newGenres = genres.filter((genre) => genre !== genreText);
      const newBlock = {
        ...info,
        blockInfo: { ...info.blockInfo, genres: newGenres },
      };
      setValue(`collectedData.${index}`, newBlock);
    }
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
        onClick={() => handleDeleteBlock(blockID)}
      >
        <DeleteIcon />
      </IconButton>

      <CBBImages blockID={index} spineColor={info.blockInfo.spineColor} />

      <MyTextArea
        name="title"
        label="Title"
        type={type}
        blockID={blockID}
        value={title || ''}
      />
      {type === 'book' ? (
        <>
          <MyTextArea
            name="author"
            label="Author"
            type={type}
            blockID={blockID}
            value={info.blockInfo.author || ''}
          />
          <MyTextArea
            name="pubYear"
            label="Pub Year"
            type={type}
            blockID={blockID}
            value={info.blockInfo.pubYear || ''}
          />
          <MyTextArea
            name="pageCount"
            label="Page Count"
            type={type}
            blockID={blockID}
            value={info.blockInfo.pageCount || ''}
          />
        </>
      ) : (
        <></>
      )}
      {type === 'book' ? (
        <GenreCheckboxes
          allGenres={allGenres}
          handleGenreClick={handleGenreClick}
          blockGenres={genres}
        />
      ) : (
        <></>
      )}
    </div>
  );
});

export default CollectedCoversBlock;
