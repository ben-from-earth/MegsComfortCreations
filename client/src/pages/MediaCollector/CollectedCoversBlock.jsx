//react
import { useDispatch } from 'react-redux';
import { memo, useContext, useEffect, useState } from 'react';

//import icons and items from Material UI
import BookIcon from '@mui/icons-material/BookTwoTone';
import MovieIcon from '@mui/icons-material/LocalMoviesTwoTone';
import VideoGameIcon from '@mui/icons-material/VideogameAssetTwoTone';
import AlbumIcon from '@mui/icons-material/AlbumTwoTone';
import IconButton from '@mui/material/IconButton';
import DeleteIcon from '@mui/icons-material/Delete';

//necessary imports from database state slice
import {
  addToDatabaseData,
  populateDatabaseData,
  removeFromDatabaseData,
} from '@/state/databaseDataSlice';

//necessary imports from png state slice
import { addToPNGCollectionList } from '@/state/pngCollectionSlice';

//genres from context provider to populate genre list based on what genres are in the database
import GenreContext from '@/context/GenreContext';

//components
import GenreCheckboxes from '@/pages/MediaCollector/GenreCheckboxes';
import CCBImages from '@/pages/MediaCollector/CCBImages';
import MyTextArea from '@/components/MyTextArea';

//setup memo so block doesnt rerender during other actions
const CollectedCoversBlock = memo(function CollectedCoversBlock({
  info: {
    type,
    images,
    blockInfo: {
      title,
      author,
      pub_year,
      page_count,
      spine_color = '#ffffff',
      databaseGenres = [],
    },
    blockID,
    isDatabase,
  },
  handleDeleteBlock,
}) {
  //setup connection to redux slice
  const dispatch = useDispatch();

  //establish variables for icons
  const icons = {
    book: <BookIcon sx={{ position: 'absolute', top: '4px', left: '4px' }} />,
    movie: <MovieIcon sx={{ position: 'absolute', top: '4px', left: '4px' }} />,
    video_game: (
      <VideoGameIcon sx={{ position: 'absolute', top: '4px', left: '4px' }} />
    ),
    album: <AlbumIcon sx={{ position: 'absolute', top: '4px', left: '4px' }} />,
  };

  //set local state for spine color
  const [color, setColor] = useState(spine_color);

  //get genres for checkbox population
  const genres = useContext(GenreContext);

  //extra information used for books only
  const bookSpecificDatabasePayload = {
    author,
    pub_year,
    page_count,
    genres: [],
  };

  //set up the data paylod for population off send to database state
  const databasePayload = {
    type,
    data: {
      title,
      spine_color: color,
      blockID,
      ...(type === 'book' ? bookSpecificDatabasePayload : {}),
    },
  };

  //on mount, populate the database data (in the state) with the block information if the information is not already from the database
  //If the block is populated from the database, add the image to the png collection list
  useEffect(() => {
    if (!isDatabase) {
      dispatch(populateDatabaseData(databasePayload));
    } else {
      dispatch(addToPNGCollectionList({ type, spine_color, url: images[0] }));
    }

    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dispatch]);

  //div to pick a color for the spine.
  //this is used in png creation and is required for the database row of books, movies, and video games
  const handleColorPick = async (blockID, type) => {
    if (!window.EyeDropper) {
      console.log('EyeDropper API not supported in this browser');
      return;
    }
    const eyeDropper = new EyeDropper();
    try {
      const { sRGBHex } = await eyeDropper.open();
      const spine_color = sRGBHex;
      setColor(spine_color);
      dispatch(addToDatabaseData({ blockID, type, spine_color }));
    } catch (e) {
      console.log(e);
    }
  };

  //if genre is clicked we add it to the data associated with the block and remove if unchecked
  const handleGenreClick = (genreText, checked) => {
    const type = 'book';
    if (checked) {
      dispatch(addToDatabaseData({ blockID, type, genreText }));
    } else {
      dispatch(removeFromDatabaseData({ blockID, type, genreText }));
    }
  };

  //styling of the block itself based on type
  const typeClasses = {
    book: 'bg-[#98ab88] border-[#3d770d]',
    movie: 'bg-[#323b43] border-black text-white',
    album: 'bg-[#7fa5a3] border-[#354544]',
    video_game: 'bg-[#98ab88] border-[#4e8885]',
  };

  return (
    <div
      className={`min-w-xs relative m-2.5 flex h-fit flex-col items-center gap-2.5 rounded-lg border-2 shadow-[5px_5px_30px_rgba(0,0,0,0.3)] ${typeClasses[type]}`}
    >
      {isDatabase && (
        <p className='p-1.25 absolute right-1 top-1 m-0 rounded-sm border-2 border-black bg-gray-700 font-["Just_Another_Hand"] text-2xl tracking-wider text-white'>
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
        onClick={() =>
          handleDeleteBlock({ blockID, type, deleteBlock: true, urls: images })
        }
      >
        <DeleteIcon />
      </IconButton>
      <CCBImages
        images={images}
        isDatabase={isDatabase}
        blockID={blockID}
        type={type}
        color={color}
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
        dispatch={dispatch}
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
            value={author || ''}
          />
          <MyTextArea
            name="pub_year"
            label="Pub Year"
            type={type}
            blockID={blockID}
            value={pub_year || ''}
          />
          <MyTextArea
            name="page_count"
            label="Page Count"
            type={type}
            blockID={blockID}
            value={page_count || ''}
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
