import Button from '@/components/Button';
import { titleRearrange } from '@/pages/MediaCollector/helpers/mediaCollectorHelpers';
import AreYouSure from '@/components/AreYouSure';
import { useState } from 'react';

//axios
import axios from 'axios';

//server domain for axios requests
const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

const DatabaseItem = ({
  info: { title, author, page_count, pub_year, spine_color, image_urls },
  type,
  handleGetMedia,
}) => {
  //classes based on type
  const typeClasses = {
    book: 'bg-[#98ab88] border-[#3d770d]',
    movie: 'bg-[#323b43] border-black text-white',
    album: 'bg-[#7fa5a3] border-[#d49a97]',
    video_game: 'bg-[#98ab88] border-[#4e8885]',
  };

  const [areYouSure, setAreYouSure] = useState(false);

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

  const onEdit = (title) => {
    console.log(`Let's edit ${title}`);
  };

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
      {image_urls.map((src, idx) => (
        <img
          key={idx}
          className={type === 'album' ? 'h-[75px]' : 'w-15'}
          src={src}
        ></img>
      ))}
      {type === 'book' ? (
        <p className='font-["Just_Another_Hand"] text-2xl'>
          {titleRearrange(title)} by {author} // {page_count} pages //{' '}
          {pub_year}
        </p>
      ) : (
        <p>{titleRearrange(title)}</p>
      )}
      <div className="ml-auto flex flex-col gap-2">
        <Button
          label={'Edit'}
          width={75}
          fontSize={24}
          onClick={() => onEdit(title)}
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
};

export default DatabaseItem;
