// react
import { memo, useContext, useState } from 'react';

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
  book: 'bg-darkpink border-[#805052]',
  movie: 'bg-[#323b43] border-black text-white',
  album: 'bg-[#7fa5a3] border-[#354544]',
  videoGame: 'bg-[#98ab88] border-[#4e8885]',
  hasError: 'bg-[#E86C54] border-[#EB4423]',
};

const CollectedCoversBlock = memo(function CollectedCoversBlock({
  hasError,
  index,
  info,
  handleDeleteBlock,
}: CollectedCoversBlockProps) {
  const {
    type,
    blockInfo: { title, spineColor = '#ffffff', genres = [] },
    blockID,
    isDatabase,
  } = info;

  //set local state for spine color
  const [color, setColor] = useState(spineColor);

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
    movie: <MovieIcon sx={{ position: 'absolute', top: '4px', left: '4px' }} />,
    videoGame: (
      <VideoGameIcon sx={{ position: 'absolute', top: '4px', left: '4px' }} />
    ),
    album: <AlbumIcon sx={{ position: 'absolute', top: '4px', left: '4px' }} />,
  };

  //div to pick a color for the spine.
  //this is used in png creation and is required for each media type in the database
  const handleColorPick = async (blockID: number) => {
    if (!window.EyeDropper) {
      console.log('EyeDropper API not supported in this browser');
      return;
    }
    const eyeDropper = new window.EyeDropper();
    try {
      const { sRGBHex } = await eyeDropper.open();
      const spineColor = sRGBHex;
      setColor(spineColor);
      const newBlock = {
        ...info,
        blockInfo: { ...info.blockInfo, spineColor },
      };
      setValue(`collectedData.${blockID}`, newBlock);
    } catch (e) {
      console.log(e);
    }
  };

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
      console.log('newBlock in genre select', newBlock);
      setValue(`collectedData.${index}`, newBlock);
    }
  };

  return (
    <div
      className={`relative flex h-fit w-lg flex-col items-center gap-2.5 rounded-lg border-2 shadow-[5px_5px_30px_rgba(0,0,0,0.3)] ${hasError ? blockClasses.hasError : blockClasses[type]}`}
    >
      {isDatabase && (
        <p className="absolute top-1 right-1 m-0 rounded-sm border-2 border-black bg-gray-700 p-1.25 tracking-wider text-white">
          Database
        </p>
      )}
      <div className="absolute top-1 left-1">
        {icons[type]} block: {index + 1}
      </div>

      <IconButton
        aria-label="delete"
        sx={{
          position: 'absolute',
          bottom: '4px',
          right: '4px',
          padding: '0',
          // color: type === 'movie' ? 'white' : '',
        }}
        onClick={() => handleDeleteBlock(blockID)}
      >
        <DeleteIcon />
      </IconButton>
      <CBBImages blockID={index} />
      {/* {type !== 'album' ? ( */}
      <div
        className="h-5 w-1/2 cursor-pointer"
        style={{ backgroundColor: color }}
        onClick={() => handleColorPick(index)}
      ></div>
      {/* ) : (
        <></>
      )} */}

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
