//react
import { memo, useContext, useState } from 'react';

//import icons and items from Material UI
import BookIcon from '@mui/icons-material/BookTwoTone';
import MovieIcon from '@mui/icons-material/LocalMoviesTwoTone';
import VideoGameIcon from '@mui/icons-material/VideogameAssetTwoTone';
import AlbumIcon from '@mui/icons-material/AlbumTwoTone';

//genres from context provider to populate genre list based on what genres are in the database
import GenreContext from '@/context/GenreContext';

//components
import GenreCheckboxes from '@/pages/MediaCollector/GenreCheckboxes';
import Button from '@/components/Button';
import axios from 'axios';
import DatabasePageContext from '@/context/DatabasePageContext';

//helpers
import { titleRearrange } from '@/pages/MediaCollector/helpers/mediaCollectorHelpers';

//server domain for axios requests
const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

// setup component text area for each data field in the block. Minimal implementation of MyTextArea used in Collected Covers Block
const MyTextArea = ({ name, label, type, value, setDatabaseData }) => {
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
}) {
  const [databaseData, setDatabaseData] = useState({
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
  const { handleGetMedia } = useContext(DatabasePageContext);

  //div under the covers to pick a color for the spine.
  //this is used in png creation and is required for the database row
  const handleColorPick = async () => {
    if (!window.EyeDropper) {
      console.log('EyeDropper API not supported in this browser');
      return;
    }
    const eyeDropper = new EyeDropper();
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
  const handleGenreClick = (genreText, checked) => {
    if (checked) {
      setDatabaseGenres((prev) => [...prev, genreText]);
    } else {
      setDatabaseGenres((prev) => prev.filter((genre) => genre !== genreText));
    }
  };

  const handleEditSubmit = async () => {
    const res = await axios.put(
      `${serverDomain}/database/edit/${type}`,
      databaseData,
      { validateStatus: (status) => status < 500 },
    );

    if (!res.data.actionCompleted) {
      setEdit(false);
    }
    if (res.data.actionCompleted === true) {
      const newGenres = databaseGenres;
      const linkGenres = [];
      const unlinkGenres = [];

      //get new added genres and link them
      for (let genre of newGenres) {
        if (!initialGenres.includes(genre)) {
          linkGenres.push(genre);
        }
      }
      try {
        const genreLinkRes = await axios.post(
          `${serverDomain}/genres/addLink`,
          {
            bookID: id,
            genres: linkGenres,
          },
        );
      } catch (err) {
        console.log('Genre link error');
      }

      //remove link to any genres that were removed
      for (let genre of initialGenres) {
        if (!newGenres.includes(genre)) {
          unlinkGenres.push(genre);
        }
      }
      try {
        const genreUnlinkRes = await axios.post(
          `${serverDomain}/genres/unlink`,
          {
            bookID: id,
            genres: unlinkGenres,
          },
        );
      } catch (err) {
        console.log('Genre unlink error');
      }

      handleGetMedia();
      setEdit(false);
    }
  };

  //classes based on type
  const typeClasses = {
    book: 'bg-[#98ab88] border-[#3d770d]',
    movie: 'bg-[#323b43] border-black text-white',
    album: 'bg-[#7fa5a3] border-[#354544]',
    video_game: 'bg-[#98ab88] border-[#4e8885]',
  };

  return (
    <div className='z-100 border-3 fixed left-1/2 top-1/2 flex -translate-x-1/2 -translate-y-1/2 flex-col content-center items-center justify-center gap-1 rounded-md border-[var(--darkpink)] bg-[var(--lightpink)] p-2 font-["Just_Another_Hand"] text-2xl tracking-wider text-black'>
      <h1>Editing: {titleRearrange(title)}</h1>

      <div
        className={`relative flex h-fit w-fit flex-col items-center gap-2.5 rounded-lg border-2 ${typeClasses[type]} mb-1`}
      >
        {icons[type]}
        <div className="gap-7.5 m-2.5 mb-0 flex flex-row items-center">
          {images.map((src) => (
            <div className="relative z-10 overflow-hidden rounded-sm" key={src}>
              <img
                className={
                  type === 'album'
                    ? 'w-21 block cursor-pointer object-cover outline-2'
                    : 'w-21 h-31 block cursor-pointer'
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
        <Button label="Close" onClick={() => setEdit(false)} width={100} />
        <Button label="Submit Changes" onClick={handleEditSubmit} />
      </div>
    </div>
  );
});

export default EditDatabaseBlock;
