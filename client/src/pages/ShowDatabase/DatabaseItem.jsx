//react
import { memo, useContext, useEffect, useState } from 'react';

//helpers
import { titleRearrange } from '@/pages/MediaCollector/helpers/mediaCollectorHelpers';

//axios
import axios from 'axios';

//necessary components
import AreYouSure from '@/components/AreYouSure';
import Button from '@/components/Button';
import EditDatabaseBlock from '@/pages/ShowDatabase/EditDatabaseBlock';
import DatabasePageContext from '@/context/DatabasePageContext';

//server domain for axios requests
const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

const DatabaseItem = memo(function DatabaseItem({
  info: { id, title, author, page_count, pub_year, spine_color, image_urls },
}) {
  //classes based on type
  const typeClasses = {
    book: 'bg-[#98ab88] border-[#3d770d]',
    movie: 'bg-[#323b43] border-black text-white',
    album: 'bg-[#7fa5a3] border-[#354544]',
    video_game: 'bg-[#98ab88] border-[#4e8885]',
  };

  // get necessary information from context
  const { type, handleGetMedia } = useContext(DatabasePageContext);

  //set up local state
  const [areYouSure, setAreYouSure] = useState(false);
  const [edit, setEdit] = useState(false);
  const [genres, setGenres] = useState([]);
  const [deleteError, setDeleteError] = useState('');

  //on mount, get the genres related to the displayed book and make sure to update if the item is edited
  useEffect(() => {
    (async () => {
      const genreRes = await axios.get(
        `${serverDomain}/genres/getForBook?bookID=${id}`,
      );

      setGenres(genreRes.data.genres);
    })();
  }, [edit]);

  //handle deletion of the media from the database
  const onDelete = async () => {
    try {
      const deleteRes = await axios.delete(
        `${serverDomain}/database?title=${title}&type=${type}`,
      );

      //need to remove links to this book ID if a book is deleted
      const removeGenreLinksRes = await axios.get(
        `${serverDomain}/genres/removeAllLinksForBook?bookID=${id}`,
      );
      if (deleteRes.data.errors || removeGenreLinksRes.data.errors) {
        setDeleteError(res.data.message);
      } else {
        handleGetMedia();
      }
    } catch (error) {
      setDeleteError('There was an error deleting, try again.');
    }

    setAreYouSure(false);
  };

  //small helper for displaying list of genres
  const list = new Intl.ListFormat('en', {
    style: 'long',
    type: 'conjunction',
  });

  return (
    <div
      className={`mr-auto box-border flex w-full items-center justify-start rounded-sm border-2 p-2 ${typeClasses[type]}`}
    >
      {areYouSure && (
        <AreYouSure
          setAreYouSure={setAreYouSure}
          onDelete={onDelete}
          title={title}
          deleteError={deleteError}
        />
      )}
      {edit && (
        <EditDatabaseBlock
          info={{
            type,
            images: image_urls,
            blockInfo: {
              title,
              author,
              pub_year,
              page_count,
              spine_color,
              initialGenres: [...genres],
            },
            id,
            setEdit,
          }}
        />
      )}
      {type !== 'album' ? (
        <div
          className={`h-36 w-6 rounded-sm`}
          style={{ backgroundColor: spine_color }}
        ></div>
      ) : (
        <></>
      )}

      {image_urls.map((src, idx) => (
        <img
          key={idx}
          className={
            type === 'album'
              ? 'ml-2 mr-7 h-36 w-36 rounded-sm'
              : 'ml-2 mr-7 h-36 w-24 rounded-sm'
          }
          src={src}
        ></img>
      ))}
      {type === 'book' ? (
        <div className="flex flex-col">
          <p className='font-["Just_Another_Hand"] text-3xl'>
            {titleRearrange(title)}
          </p>
          <p className='font-["Just_Another_Hand"] text-2xl'>{author}</p>
          <hr className="my-1 border-t border-black" />
          <p className='font-["Just_Another_Hand"] text-xl'>
            Pages: {page_count}
          </p>

          <p className='font-["Just_Another_Hand"] text-xl'>
            Publication Date: {pub_year}
          </p>
          <p className='font-["Just_Another_Hand"] text-xl'>
            Genres: {list.format(genres)}
          </p>
        </div>
      ) : (
        <p className='font-["Just_Another_Hand"] text-4xl'>
          {titleRearrange(title)}
        </p>
      )}
      <div className="ml-auto flex flex-col gap-2">
        <Button
          label={'Edit'}
          width={75}
          fontSize={24}
          onClick={() => setEdit(true)}
        />
        <Button
          label={'Delete'}
          width={75}
          fontSize={24}
          onClick={() => setAreYouSure(true)}
        />
      </div>
    </div>
  );
});

export default DatabaseItem;
