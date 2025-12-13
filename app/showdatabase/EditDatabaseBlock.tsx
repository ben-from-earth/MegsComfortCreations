// react, redux imports
import { Dispatch, memo, SetStateAction, useContext, useState } from 'react';

//import icons and items from Material UI
import BookIcon from '@mui/icons-material/BookTwoTone';
import MovieIcon from '@mui/icons-material/LocalMoviesTwoTone';
import VideoGameIcon from '@mui/icons-material/VideogameAssetTwoTone';
import AlbumIcon from '@mui/icons-material/AlbumTwoTone';

// context
import GenreContext from '@/lib/context/GenreContext';
import { useDatabasePageContext } from '@/lib/context/DatabasePageContext';

// components
import GenreCheckboxes from '@/app/mediacollector/GenreCheckboxes';
import Button from '@/app/components/Button';

// library imports
import axios from 'axios';

// helpers
import { titleRearrange } from '@/lib/helpers/titleRearrange';
import { mediaTypeBlockClasses } from '@/app/mediacollector/CollectedCoversBlock';

// interfaces and types
import {
  blockInfo,
  MediaType,
  postSavedMediaItem,
  SuccessfulMediaSaveEditResponse,
} from '@/lib/interfaces/globalInterfaces';
import {
  DatabaseSaveEditErrorResponse,
  ErrorResponse,
} from '@/app/api/api-Errors';

export interface MinimalTextAreaProps {
  name: 'title' | 'author' | 'pub_year' | 'page_count';
  label: string;
  type: MediaType;
  value: string | number;
  setDatabaseData: Dispatch<SetStateAction<postSavedMediaItem>>;
}

export interface EditDatabaseBlockProps {
  info: {
    type: MediaType;
    images: string[];
    blockInfo: Omit<blockInfo, 'databaseGenres'> & { initialGenres: string[] };
    id: string;
    setEdit: Dispatch<SetStateAction<boolean>>;
  };
}

// setup component text area for each data field in the block. Minimal implementation of MyTextArea used in Collected Covers Block
const MyTextArea = ({
  name,
  label,
  type,
  value,
  setDatabaseData,
}: MinimalTextAreaProps) => {
  const labelClass =
    type === 'book'
      ? 'w-25 content-center text-right text-2xl'
      : 'w-15 content-center text-right text-2xl';

  return (
    <div className="grid grid-cols-[max-content_1fr] gap-x-3 gap-y-1 p-2">
      <label className={labelClass} htmlFor={name}>
        {label}:
      </label>
      <textarea
        className="w-2xs content-center rounded-sm bg-white pl-2 text-black"
        name={name}
        defaultValue={value}
        onChange={(e) => {
          if (name === 'pub_year' || name === 'page_count') {
            setDatabaseData((prev) => ({
              ...prev,
              [name]: Number(e.target.value),
            }));
          } else {
            setDatabaseData((prev) => ({ ...prev, [name]: e.target.value }));
          }
        }}
      ></textarea>
    </div>
  );
};

//setup memo so block doesnt rerender during other actions
const EditDatabaseBlock = memo(function EditDatabaseBlock({
  info: {
    type,
    images,
    blockInfo: {
      title,
      author,
      pub_year,
      page_count,
      spine_color = '#ffffff',
      initialGenres = [],
    },
    id,
    setEdit,
  },
}: EditDatabaseBlockProps) {
  const [databaseData, setDatabaseData] = useState<postSavedMediaItem>({
    id,
    title,
    author,
    pub_year,
    page_count,
    spine_color,
    image_urls: images,
  });

  const [databaseGenres, setDatabaseGenres] = useState([...initialGenres]);
  const [color, setColor] = useState(spine_color);

  //establish variables for icons
  const icons = {
    book: <BookIcon sx={{ position: 'absolute', top: '4px', left: '4px' }} />,
    movie: <MovieIcon sx={{ position: 'absolute', top: '4px', left: '4px' }} />,
    video_game: (
      <VideoGameIcon sx={{ position: 'absolute', top: '4px', left: '4px' }} />
    ),
    album: <AlbumIcon sx={{ position: 'absolute', top: '4px', left: '4px' }} />,
  };

  //get genres for checkbox population
  const genres = useContext(GenreContext);
  const { handleGetMedia } = useDatabasePageContext();

  //div under the covers to pick a color for the spine.
  //this is used in png creation and is required for the database row
  const handleColorPick = async () => {
    if (!window.EyeDropper) {
      console.log('EyeDropper API not supported in this browser');
      return;
    }
    const eyeDropper = new window.EyeDropper();
    try {
      const { sRGBHex } = await eyeDropper.open();
      const spine_color = sRGBHex;
      setColor(spine_color);
      setDatabaseData((prev) => ({ ...prev, spine_color: spine_color }));
    } catch (e) {
      console.log(e);
    }
  };

  //if genre is clicked we add it to the data associated with the block and remove if unchecked
  const handleGenreClick = (genreText: string, checked: boolean) => {
    if (checked) {
      setDatabaseGenres((prev) => [...prev, genreText]);
    } else {
      setDatabaseGenres((prev) => prev.filter((genre) => genre !== genreText));
    }
  };

  const handleEditSubmit = async () => {
    const res = await axios.put<
      | ErrorResponse
      | SuccessfulMediaSaveEditResponse
      | DatabaseSaveEditErrorResponse
    >(`/api/database/edit/${type}`, databaseData, {
      validateStatus: (status) => status < 500,
    });

    if ('error' in res.data) {
      // just closing the window if there are any errors
      // need to display error message
      setEdit(false);
    } else {
      const newGenres = databaseGenres;
      const linkGenres: string[] = [];
      const unlinkGenres: string[] = [];

      //get new added genres and link them
      for (const genre of newGenres) {
        if (!initialGenres.includes(genre)) {
          linkGenres.push(genre);
        }
      }
      try {
        await axios.post(`/api/genres/addlink`, {
          bookID: id,
          genres: linkGenres,
        });
      } catch {
        console.log('Genre link error');
      }

      //remove link to any genres that were removed
      for (const genre of initialGenres) {
        if (!newGenres.includes(genre)) {
          unlinkGenres.push(genre);
        }
      }
      try {
        await axios.post(`/api/genres/unlink`, {
          bookID: id,
          genres: unlinkGenres,
        });
      } catch {
        console.log('Genre unlink error');
      }

      handleGetMedia();
      setEdit(false);
    }
  };

  return (
    <div className="border-darkpink bg-lightpink fixed top-1/2 left-1/2 z-100 flex -translate-x-1/2 -translate-y-1/2 flex-col content-center items-center justify-center gap-1 rounded-md border-3 p-2 text-2xl tracking-wider text-black">
      <h1>Editing: {titleRearrange(title)}</h1>

      <div
        className={`relative flex h-fit w-fit flex-col items-center gap-2.5 rounded-lg border-2 text-lg ${mediaTypeBlockClasses[type]} mb-1`}
      >
        {icons[type]}
        <div className="m-2.5 mb-0 flex flex-row items-center gap-7.5">
          {images.map((src) => (
            <div className="relative z-10 overflow-hidden rounded-sm" key={src}>
              <img
                className={
                  type === 'album'
                    ? 'block w-21 cursor-pointer object-cover outline-2'
                    : 'block h-31 w-21 cursor-pointer'
                }
                src={src}
              ></img>
            </div>
          ))}
        </div>
        {type !== 'album' ? (
          <div
            className="h-5 w-1/2 cursor-pointer"
            style={{ backgroundColor: color }}
            onClick={() => handleColorPick()}
          ></div>
        ) : (
          <></>
        )}

        <MyTextArea
          name="title"
          label="Title"
          type={type}
          value={titleRearrange(title) || ''}
          setDatabaseData={setDatabaseData}
        />
        {type === 'book' ? (
          <>
            <MyTextArea
              name="author"
              label="Author"
              type={type}
              value={author || ''}
              setDatabaseData={setDatabaseData}
            />
            <MyTextArea
              name="pub_year"
              label="Pub Year"
              type={type}
              value={pub_year || ''}
              setDatabaseData={setDatabaseData}
            />
            <MyTextArea
              name="page_count"
              label="Page Count"
              type={type}
              value={page_count || ''}
              setDatabaseData={setDatabaseData}
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
      <div className="flex gap-2">
        <Button
          label="Close"
          onClick={() => setEdit(false)}
          width={100}
          fontSize={25}
        />
        <Button
          label="Submit Changes"
          onClick={handleEditSubmit}
          width={150}
          fontSize={25}
        />
      </div>
    </div>
  );
});

export default EditDatabaseBlock;
