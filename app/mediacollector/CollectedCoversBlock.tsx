// react
import { memo, useContext, useEffect, useState } from 'react';
import { useAppDispatch } from 'lib/state/store';

//import icons and items from Material UI
import BookIcon from '@mui/icons-material/BookTwoTone';
import MovieIcon from '@mui/icons-material/LocalMoviesTwoTone';
import VideoGameIcon from '@mui/icons-material/VideogameAssetTwoTone';
import AlbumIcon from '@mui/icons-material/AlbumTwoTone';
import IconButton from '@mui/material/IconButton';
import DeleteIcon from '@mui/icons-material/Delete';

// necessary imports from database state slice
import {
  addToDatabaseData,
  populateDatabaseData,
  removeFromDatabaseData,
} from 'lib/state/slices/databaseDataSlice';

// necessary imports from png state slice
import { addToPNGCollectionList } from 'lib/state/slices/pngCollectionSlice';

// components
import CBBImages from '@//mediacollector/CBBImages';
import MyTextArea from '@/shared/MyTextArea';
import GenreContext from 'lib/context/GenreContext';
import GenreCheckboxes from '@//mediacollector/GenreCheckboxes';

// interfaces and types
import {
  AlbumInsert,
  BookInsert,
  MediaType,
  MovieInsert,
  VideoGameInsert,
} from 'lib/interfaces/globalInterfaces';
import { CollectedBlockInformation } from 'lib/state/slices/collectorSlice';

export type DatabasePayload =
  | { type: 'book'; data: BookInsert }
  | { type: 'movie'; data: MovieInsert }
  | { type: 'videoGame'; data: VideoGameInsert }
  | { type: 'album'; data: AlbumInsert };

export interface CollectedCoversBlockProps {
  info: CollectedBlockInformation;
  handleDeleteBlock: (
    blockID: string,
    type: MediaType,
    deleteBlock: boolean,
    urls: string[],
  ) => void;
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
export const mediaTypeBlockClasses = {
  book: 'bg-darkpink border-[#805052]',
  movie: 'bg-[#323b43] border-black text-white',
  album: 'bg-[#7fa5a3] border-[#354544]',
  videoGame: 'bg-[#98ab88] border-[#4e8885]',
};

const CollectedCoversBlock = memo(function CollectedCoversBlock({
  info,
  handleDeleteBlock,
}: CollectedCoversBlockProps) {
  const {
    type,
    images,
    blockInfo: { title, spineColor = '#ffffff', databaseGenres = [] },
    blockID,
    isDatabase,
  } = info;

  //setup connection to redux slice
  const dispatch = useAppDispatch();

  //establish variables for icons
  const icons = {
    book: <BookIcon sx={{ position: 'absolute', top: '4px', left: '4px' }} />,
    movie: <MovieIcon sx={{ position: 'absolute', top: '4px', left: '4px' }} />,
    videoGame: (
      <VideoGameIcon sx={{ position: 'absolute', top: '4px', left: '4px' }} />
    ),
    album: <AlbumIcon sx={{ position: 'absolute', top: '4px', left: '4px' }} />,
  };

  //set local state for spine color
  const [color, setColor] = useState(spineColor);

  //get genres for checkbox population
  const genres = useContext(GenreContext);

  //extra information used for books only
  const baseData = {
    title,
    spineColor: color,
    blockID,
    imageUrls: [] as string[],
  };

  let databasePayload: DatabasePayload;

  if (type === 'book') {
    const {
      blockInfo: { author, pubYear, pageCount },
    } = info;
    databasePayload = {
      type: 'book',
      data: {
        ...baseData,
        author: author ?? '',
        pubYear: pubYear ?? null,
        pageCount: pageCount ?? null,
        genres: [],
      },
    };
  } else if (type === 'movie') {
    databasePayload = {
      type: 'movie',
      data: {
        ...baseData,
      },
    };
  } else if (type === 'videoGame') {
    databasePayload = {
      type: 'videoGame',
      data: {
        ...baseData,
      },
    };
  } else {
    databasePayload = {
      type: 'album',
      data: {
        ...baseData,
      },
    };
  }

  //on mount, populate the database data (in the state) with the block information if the information is not already from the database
  //If the block is populated from the database, add the image to the png collection list
  useEffect(() => {
    if (!isDatabase) {
      dispatch(populateDatabaseData(databasePayload));
    } else {
      dispatch(addToPNGCollectionList({ type, spineColor, url: images[0] }));
    }

    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dispatch]);

  //div to pick a color for the spine.
  //this is used in png creation and is required for each media type in the database
  const handleColorPick = async (blockID: string, type: MediaType) => {
    if (!window.EyeDropper) {
      console.log('EyeDropper API not supported in this browser');
      return;
    }
    const eyeDropper = new window.EyeDropper();
    try {
      const { sRGBHex } = await eyeDropper.open();
      const spineColor = sRGBHex;
      setColor(spineColor);
      dispatch(addToDatabaseData({ blockID, type, spineColor }));
    } catch (e) {
      console.log(e);
    }
  };

  //if genre is clicked we add it to the data associated with the block and remove if unchecked
  const handleGenreClick = (genreText: string, checked: boolean) => {
    if (checked) {
      dispatch(addToDatabaseData({ blockID, type: 'book', genreText }));
    } else {
      dispatch(removeFromDatabaseData({ blockID, type: 'book', genreText }));
    }
  };

  return (
    <div
      className={`relative flex h-fit min-w-sm flex-col items-center gap-2.5 rounded-lg border-2 shadow-[5px_5px_30px_rgba(0,0,0,0.3)] ${mediaTypeBlockClasses[type]}`}
    >
      {isDatabase && (
        <p className="absolute top-1 right-1 m-0 rounded-sm border-2 border-black bg-gray-700 p-1.25 tracking-wider text-white">
          Database
        </p>
      )}
      {icons[type]}
      <IconButton
        aria-label="delete"
        sx={{
          position: 'absolute',
          bottom: '4px',
          right: '4px',
          padding: '0',
          color: type === 'movie' ? 'white' : '',
        }}
        onClick={() => handleDeleteBlock(blockID, type, true, images)}
      >
        <DeleteIcon />
      </IconButton>
      <CBBImages
        images={images}
        isDatabase={isDatabase}
        blockID={blockID}
        type={type}
        spineColor={color}
      />
      {type !== 'album' ? (
        <div
          className="h-5 w-1/2 cursor-pointer"
          style={{ backgroundColor: color }}
          onClick={() => handleColorPick(blockID, type)}
        ></div>
      ) : (
        <></>
      )}

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
          genres={genres}
          databaseGenres={databaseGenres}
          handleGenreClick={handleGenreClick}
        />
      ) : (
        <></>
      )}
    </div>
  );
});

export default CollectedCoversBlock;
