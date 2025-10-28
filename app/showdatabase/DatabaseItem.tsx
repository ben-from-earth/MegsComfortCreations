'use client';
// react, redux imports
import { memo, useEffect, useState } from 'react';

// library imports
import axios from 'axios';

// components
import AreYouSure from '@/app/components/AreYouSure';
import Button from '@/app/components/Button';
import { mediaTypeBlockClasses } from '@/app/mediacollector/CollectedCoversBlock';
import EditDatabaseBlock from '@/app/showdatabase/EditDatabaseBlock';

// helpers
import { titleRearrange } from '@/lib/helpers/titleRearrange';

// context
import { useDatabasePageContext } from '@/lib/context/DatabasePageContext';

// interfaces and types
import { ErrorResponse } from '@/app/api/api-Errors';
import { postSavedMediaItem } from '@/lib/interfaces/globalInterfaces';

export interface DatabaseItemProps {
  info: postSavedMediaItem;
}

const DatabaseItem = memo(function DatabaseItem({
  info: { id, title, author, page_count, pub_year, spine_color, image_urls },
}: DatabaseItemProps) {
  // get necessary information from context
  const { type, handleGetMedia } = useDatabasePageContext();

  //set up local state
  const [areYouSure, setAreYouSure] = useState(false);
  const [edit, setEdit] = useState(false);
  const [genres, setGenres] = useState<string[]>([]);
  //   const [deleteError, setDeleteError] = useState<string | undefined>();

  //on mount, get the genres related to the displayed book and make sure to update if the item is edited
  useEffect(() => {
    (async () => {
      const genreRes = await axios.get<
        { message: string; genres: string[] } | ErrorResponse
      >(`api/genres/getforbook?bookID=${id}`);

      if ('error' in genreRes.data === false) setGenres(genreRes.data.genres);
    })();
  }, [edit]);

  //handle deletion of the media from the database
  const onDelete = async () => {
    try {
      const deleteRes = await axios.delete<ErrorResponse | { message: string }>(
        `api/database/delete?title=${title}&type=${type}`,
      );
      const response = deleteRes.data;
      if ('error' in response) {
        console.log(response.message);
        // setDeleteError(response.message);
      } else {
        handleGetMedia();
      }
    } catch (error) {
      console.log('There was an error deleting, try again.');
      //   setDeleteError('There was an error deleting, try again.');
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
      className={`mr-auto box-border flex w-full items-center justify-start rounded-sm border-2 p-2 ${mediaTypeBlockClasses[type]}`}
    >
      {areYouSure && (
        <AreYouSure
          setAreYouSure={setAreYouSure}
          onDelete={onDelete}
          title={title}
          //   deleteError={deleteError}
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
              ? 'mr-7 h-36 w-36 rounded-sm'
              : 'mr-7 ml-2 h-36 w-24 rounded-sm'
          }
          src={src}
        ></img>
      ))}
      {type === 'book' ? (
        <div className="flex flex-col">
          <p className="text-3xl">{titleRearrange(title)}</p>
          <p className="text-2xl">{author}</p>
          <hr className="my-1 border-t border-black" />
          <p className="text-xl">Pages: {page_count}</p>

          <p className="text-xl">Publication Date: {pub_year}</p>
          <p className="text-xl">Genres: {list.format(genres)}</p>
        </div>
      ) : (
        <p className="text-4xl">{titleRearrange(title)}</p>
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
