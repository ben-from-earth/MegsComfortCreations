import Button from '@/components/Button';
import { titleRearrange } from '@/pages/MediaCollector/helpers/mediaCollectorHelpers';
import AreYouSure from '@/components/AreYouSure';
import { memo, useEffect, useState } from 'react';

//axios
import axios from 'axios';
import EditDatabaseBlock from '@/pages/ShowDatabase/EditDatabaseBlock';

//server domain for axios requests
const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

const DatabaseItem = memo(function DatabaseItem({
  info: { id, title, author, page_count, pub_year, spine_color, image_urls },
  type,
  handleGetMedia,
}) {
  //classes based on type
  const typeClasses = {
    book: 'bg-[#98ab88] border-[#3d770d]',
    movie: 'bg-[#323b43] border-black text-white',
    album: 'bg-[#7fa5a3] border-[#d49a97]',
    video_game: 'bg-[#98ab88] border-[#4e8885]',
  };

  const [areYouSure, setAreYouSure] = useState(false);
  const [edit, setEdit] = useState(false);
  const [genres, setGenres] = useState([]);

  useEffect(() => {
    (async () => {
      const genreRes = await axios.post(`${serverDomain}/genres/getFromBook`, {
        bookID: id,
      });

      setGenres(genreRes.data.genres);
    })();
  }, [edit]);

  const onDelete = async () => {
    try {
      const res = await axios.delete(
        `${serverDomain}/database?title=${title}&type=${type}`,
      );
      if (res.data.errors) setDeleteError(res.data.message);
      handleGetMedia();
    } catch (error) {
      setDeleteError('There was an error deleting, try again.');
    }

    setAreYouSure(false);
  };

  const list = new Intl.ListFormat('en', {
    style: 'long',
    type: 'conjunction',
  });

  return (
    <div
      className={`mr-auto box-border flex w-full items-center justify-start gap-5 rounded-sm border-2 p-2 ${typeClasses[type]}`}
    >
      {areYouSure && (
        <AreYouSure
          setAreYouSure={setAreYouSure}
          onDelete={onDelete}
          title={title}
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
      {image_urls.map((src, idx) => (
        <img
          key={idx}
          className={type === 'album' ? 'h-30 rounded-sm' : 'w-24 rounded-sm'}
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
        <p className='font-["Just_Another_Hand"] text-5xl'>
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
