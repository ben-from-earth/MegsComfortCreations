//import icons and items from Material UI
import BookIcon from '@mui/icons-material/BookTwoTone';
import MovieIcon from '@mui/icons-material/LocalMoviesTwoTone';
import VideoGameIcon from '@mui/icons-material/VideogameAssetTwoTone';
import AlbumIcon from '@mui/icons-material/AlbumTwoTone';
import IconButton from '@mui/material/IconButton';
import DeleteIcon from '@mui/icons-material/Delete';

//react imports
import { useDispatch } from 'react-redux';
import { memo, useContext, useEffect, useState } from 'react';

//necessary imports from database state slice
import {
  addToDatabaseData,
  populateDatabaseData,
  removeFromDatabaseData,
  updateDatabaseData,
} from '@/state/databaseDataSlice';

//necessary imports from png state slice
import {
  addToPNGCollectionList,
  removeFromPNGCollectionList,
} from '@/state/pngCollectionSlice';

//genres from context provider to populate genre list based on what genres are in the database
import GenreContext from '@/context/GenreContext';

//components
import GenreCheckboxes from './GenreCheckboxes';

// setup component text area for each data field in the block
const MyTextArea = ({ name, label, type, blockID, value }) => {
  const dispatch = useDispatch();

  const labelClass =
    type === 'book'
      ? 'w-25 content-center text-right font-["Just_Another_Hand"] text-3xl'
      : 'w-15 content-center text-right font-["Just_Another_Hand"] text-3xl';

  return (
    <div className="grid grid-cols-[max-content_1fr] gap-x-3 gap-y-1 p-2">
      <label className={labelClass} htmlFor={name}>
        {label}:
      </label>
      <textarea
        className="w-full content-center whitespace-nowrap rounded-sm bg-white pl-2 text-black"
        name={name}
        defaultValue={value}
        onChange={(e) => {
          dispatch(
            updateDatabaseData({
              blockID,
              type,
              name,
              newText: e.target.value,
            }),
          );
        }}
      ></textarea>
    </div>
  );
};

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
    book: (
      <BookIcon sx={{ position: 'absolute', bottom: '4px', left: '4px' }} />
    ),
    movie: (
      <MovieIcon sx={{ position: 'absolute', bottom: '4px', left: '4px' }} />
    ),
    video_game: (
      <VideoGameIcon
        sx={{ position: 'absolute', bottom: '4px', left: '4px' }}
      />
    ),
    album: (
      <AlbumIcon sx={{ position: 'absolute', bottom: '4px', left: '4px' }} />
    ),
  };

  //set local state for spine color
  const [color, setColor] = useState(spine_color);

  //set up a local state to an array with an index for each image slot (current: 3) and set to false
  //this is for click tracking and the "selected" style and setting the image as a block datapoint
  const [clicked, setClicked] = useState(() =>
    Array(images.length).fill(false),
  );

  //get genres for checkbox population
  const genres = useContext(GenreContext);

  //populated database add with following information added for book
  const bookSpecificDatabasePayload = {
    author,
    pub_year,
    page_count,
    genres: [],
  };

  const databasePayload = {
    type,
    data: {
      title,
      spine_color: color,
      blockID,
      ...(type === 'book' ? bookSpecificDatabasePayload : {}),
    },
  };

  //on mount, populate the database data (in the state) with the block information
  useEffect(() => {
    if (!isDatabase) {
      dispatch(populateDatabaseData(databasePayload));
    } else {
      dispatch(addToPNGCollectionList({ type, spine_color, url: images[0] }));
    }

    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dispatch]);

  // function for handling an image click
  //this adds the image url to the database data (in the state) or removes it if its there already
  const handleClick = (blockID, type, idx, src) => {
    const next = !clicked[idx];
    setClicked((prev) =>
      prev.map((b, itemIndex) => (itemIndex === idx ? next : b)),
    );
    if (next) {
      dispatch(
        addToDatabaseData({
          type,
          src,
          idx,
          blockID,
        }),
      );
      dispatch(addToPNGCollectionList({ url: src, type, spine_color: color }));
    } else {
      dispatch(removeFromDatabaseData({ blockID, type, idx }));
      dispatch(removeFromPNGCollectionList({ url: src }));
    }
  };

  //div under the covers to pick a color for the spine.
  //this is used in png creation and is required for the database row
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

  //classes based on type
  const typeClasses = {
    book: 'bg-[#98ab88] border-[#3d770d]',
    movie: 'bg-[#323b43] border-black text-white',
    album: 'bg-[#7fa5a3] border-[#d49a97]',
    video_game: 'bg-[#98ab88] border-[#4e8885]',
  };

  return (
    <div
      className={`relative m-2.5 flex h-fit w-fit flex-col items-center gap-2.5 rounded-lg border-2 shadow-[5px_5px_30px_rgba(0,0,0,0.3)] ${typeClasses[type]}`}
    >
      {isDatabase && (
        <p className='p-1.25 absolute right-1 top-1 m-0 rounded-sm border-2 border-black bg-gray-700 font-["Just_Another_Hand"] text-2xl tracking-wider text-white'>
          Database
        </p>
      )}
      {icons[type]}
      <IconButton
        aria-label="delete"
        sx={{ position: 'absolute', bottom: '4px', right: '4px', padding: '0' }}
        onClick={() =>
          handleDeleteBlock({ blockID, type, deleteBlock: true, urls: images })
        }
      >
        <DeleteIcon />
      </IconButton>
      <div className="gap-7.5 m-2.5 mb-0 flex flex-row items-center">
        {images.map((src, idx) => (
          <div
            className="relative z-10 overflow-hidden rounded-sm"
            key={src}
            onClick={() => {
              if (!isDatabase) {
                handleClick(blockID, type, idx, src);
              }
            }}
          >
            <img
              className={
                type === 'album'
                  ? 'w-21 block cursor-pointer object-cover outline-2'
                  : 'w-21 h-31 block cursor-pointer'
              }
              src={src}
            ></img>
            <div
              className={`pointer-events-none absolute inset-0 flex content-center items-center bg-[rgba(0,200,0,0.5)] ${clicked[idx] ? 'opacity-100' : 'opacity-0'}`}
            >
              <p className='-rotate-65 -translate-x-1 font-["Just_Another_Hand"] text-5xl font-bold tracking-wider text-[rgb(0,77,0)]'>
                Selected
              </p>
            </div>
          </div>
        ))}
      </div>
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
          blockID={blockID}
        />
      ) : (
        <></>
      )}
    </div>
  );
});

export default CollectedCoversBlock;
