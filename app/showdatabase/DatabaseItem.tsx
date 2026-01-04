'use client';
// react, redux imports
import { memo, useEffect, useState } from 'react';

// library imports
import { trpc } from 'lib/trpc/client';

// components
import AreYouSure from '@/shared/AreYouSure';
import Button from '@/shared/Button';
import { mediaTypeBlockClasses } from '@//mediacollector/CollectedCoversBlock';
import EditDatabaseBlock from '@//showdatabase/EditDatabaseBlock';

// helpers
import { titleRearrange } from 'lib/helpers/titleRearrange';

// context
import { useDatabasePageContext } from 'lib/context/DatabasePageContext';

// interfaces and types
import { PostSavedMediaItem } from 'lib/interfaces/globalInterfaces';
import { isBookRow } from 'lib/helpers/handleMediaTyping';

export interface DatabaseItemProps {
  info: PostSavedMediaItem;
}

const DatabaseItem = memo(function DatabaseItem({ info }: DatabaseItemProps) {
  // get necessary information from context
  const { type, handleGetMedia } = useDatabasePageContext();

  const { id, title, spineColor, imageUrls } = info;

  //set up local state
  const [areYouSure, setAreYouSure] = useState(false);
  const [edit, setEdit] = useState(false);
  const [genres, setGenres] = useState<string[]>([]);
  //   const [deleteError, setDeleteError] = useState<string | undefined>();

  //on mount, get the genres related to the displayed book and make sure to update if the item is edited
  const genresQuery = trpc.genres.getForBook.useQuery({ bookID: id });
  useEffect(() => {
    if (genresQuery.data?.genres) setGenres(genresQuery.data.genres);
  }, [genresQuery.data]);

  //handle deletion of the media from the database
  const { mutateAsync: databaseDelete } =
    trpc.database.deleteByTitle.useMutation();
  const onDelete = async () => {
    try {
      await databaseDelete({ title, type });
      handleGetMedia();
    } catch {
      console.log('There was an error deleting, try again.');
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
          info={
            isBookRow(type, info)
              ? {
                  type,
                  images: imageUrls ?? [],
                  BlockInfo: {
                    title,
                    author: info.author,
                    pubYear: info.pubYear,
                    pageCount: info.pageCount,
                    spineColor,
                    initialGenres: [...genres],
                  },
                  id,
                  setEdit,
                }
              : {
                  type,
                  images: imageUrls ?? [],
                  BlockInfo: {
                    title,
                    spineColor,
                    initialGenres: [...genres],
                  },
                  id,
                  setEdit,
                }
          }
        />
      )}
      {type !== 'album' ? (
        <div
          className={`h-36 w-6 rounded-sm`}
          style={{ backgroundColor: spineColor }}
        ></div>
      ) : (
        <></>
      )}

      {imageUrls?.map((src, idx) => (
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
      {type === 'book' && isBookRow(type, info) ? (
        <div className="flex flex-col">
          <p className="text-3xl">{titleRearrange(title)}</p>
          <p className="text-2xl">{info.author}</p>
          <hr className="my-1 border-t border-black" />
          <p className="text-xl">Pages: {info.pageCount}</p>

          <p className="text-xl">Publication Date: {info.pubYear}</p>
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
