'use client';
// react, redux imports
import { memo, useEffect, useState } from 'react';

// library imports
import { trpc } from 'lib/trpc/client';

// components
import AreYouSure from '@/shared/AreYouSure';
import Button from '@/components/ui/Button';
import EditDatabaseBlock from '@/showdatabase/EditDatabaseBlock';
import MediaImageStrip from '@/shared/MediaImageStrip';

// helpers
import { titleRearrange } from 'lib/helpers/titleRearrange';

// context
import { useDatabasePageContext } from 'lib/context/DatabasePageContext';

// interfaces and types
import { PostSavedMediaItem } from 'lib/interfaces/globalInterfaces';
import { isBookRow } from 'lib/helpers/handleMediaTyping';
import { MEDIA_TYPES } from 'lib/constants/mediaTypes';
import { blockClasses } from 'lib/constants/typeBlockStyles';
import { convertMediaItemToForm } from '@/mediacollector/collector-form/mediaItemFormSchema';

export interface DatabaseItemProps {
  info: PostSavedMediaItem;
}

const DatabaseItem = memo(function DatabaseItem({ info }: DatabaseItemProps) {
  // get necessary information from context
  const {
    databaseItems: { type: displayedListType },
    handleGetMedia,
  } = useDatabasePageContext();

  const { id, title, spineColor, images } = info;
  const itemType =
    MEDIA_TYPES.find((mediaType) => mediaType === info.mediaType) ??
    displayedListType;

  //set up local state
  const [areYouSure, setAreYouSure] = useState(false);
  const [edit, setEdit] = useState(false);
  const [genres, setGenres] = useState<string[]>([]);
  //   const [deleteError, setDeleteError] = useState<string | undefined>();

  //on mount, get the genres related to the displayed book and make sure to update if the item is edited
  const genresQuery = trpc.genres.getForBook.useQuery(
    { bookID: id },
    { enabled: itemType === 'book' },
  );
  useEffect(() => {
    if (genresQuery.data?.genres) setGenres(genresQuery.data.genres);
  }, [genresQuery.data]);

  //handle deletion of the media from the database
  const { mutateAsync: databaseDelete } = trpc.database.delete.useMutation();
  const onDelete = async () => {
    try {
      await databaseDelete({ id, type: itemType });
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
  const bookDetails = isBookRow(info) ? info : null;

  const mediaItem = convertMediaItemToForm({
    item: info,
    type: itemType,
    genres,
  });

  return (
    <div
      className={`mr-auto box-border flex w-full items-center justify-start rounded-sm p-2 ${blockClasses[itemType]}`}
    >
      {areYouSure && (
        <AreYouSure
          setAreYouSure={setAreYouSure}
          onDelete={onDelete}
          title={title}
        />
      )}
      {edit && <EditDatabaseBlock item={mediaItem} setEdit={setEdit} />}
      {itemType !== 'album' ? (
        <div
          className={`h-36 w-6 rounded-sm`}
          style={{ backgroundColor: spineColor }}
        ></div>
      ) : (
        <></>
      )}

      <MediaImageStrip
        mediaType={itemType}
        images={images ?? []}
        className="mr-7 ml-2 flex flex-row items-center gap-3"
        albumTileClassName="relative z-10 h-36 w-24 overflow-hidden rounded-sm"
        defaultTileClassName="relative z-10 h-36 w-24 overflow-hidden rounded-sm"
      />
      {itemType === 'book' && bookDetails ? (
        <div className="flex flex-col">
          <p className="text-3xl">{titleRearrange(title)}</p>
          <p className="text-2xl">{bookDetails.author}</p>
          <hr className="my-1 border-t border-black" />
          <p className="text-xl">Pages: {bookDetails.pageCount}</p>

          <p className="text-xl">Publication Date: {bookDetails.pubYear}</p>
          <p className="text-xl">Genres: {list.format(genres)}</p>
        </div>
      ) : (
        <p className="text-4xl">{titleRearrange(title)}</p>
      )}
      <div className="ml-auto flex flex-col gap-2">
        <Button
          variant="primary"
          label={'Edit'}
          width={75}
          fontSize={24}
          onClick={() => setEdit(true)}
        />
        <Button
          variant="primary"
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
