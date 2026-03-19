// react, redux imports
import {
  ChangeEvent,
  Dispatch,
  memo,
  SetStateAction,
  useContext,
  useRef,
  useState,
} from 'react';

// context
import GenreContext from 'lib/context/GenreContext';
import { useDatabasePageContext } from 'lib/context/DatabasePageContext';

// components
import GenreCheckboxes from '@/mediacollector/GenreCheckboxes';
import Button from '@/components/ui/Button';
import MediaImageStrip from '@/shared/MediaImageStrip';

// library imports
import { trpc } from 'lib/trpc/client';

// helpers
import { titleRearrange } from 'lib/helpers/titleRearrange';

// interfaces and types
import { BlockInfo, PostSavedMediaItem } from 'lib/interfaces/globalInterfaces';
import { MediaType } from 'lib/constants/mediaTypes';
import { blockClasses, icons } from 'lib/constants/typeBlockStyles';

export interface MinimalTextAreaProps {
  name: 'title' | 'author' | 'pubYear' | 'pageCount';
  label: string;
  type: MediaType;
  value: string | number;
  setDatabaseData: Dispatch<SetStateAction<PostSavedMediaItem>>;
}

export interface EditDatabaseBlockProps {
  info: {
    type: MediaType;
    images: PostSavedMediaItem['images'];
    blockInfo: Omit<BlockInfo, 'databaseGenres'> & { initialGenres: string[] };
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
          if (name === 'pubYear' || name === 'pageCount') {
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
      pubYear,
      pageCount,
      spineColor = '#ffffff',
      initialGenres = [],
    },
    id,
    setEdit,
  },
}: EditDatabaseBlockProps) {
  const initialValues = {
    id,
    title,
    author: author ?? '',
    pubYear,
    pageCount,
    spineColor,
    images:
      images.length > 0
        ? images
        : [
            {
              url: '/images/placeholder-image.png',
              isDefault: true,
              selected: true,
              spineColor,
            },
          ],
  };

  const [databaseData, setDatabaseData] =
    useState<PostSavedMediaItem>(initialValues);

  const [databaseGenres, setDatabaseGenres] = useState([...initialGenres]);
  const [color, setColor] = useState(spineColor);
  const [isUploading, setIsUploading] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  //get genres for checkbox population
  const genres = useContext(GenreContext);
  const { handleGetMedia } = useDatabasePageContext();

  const defaultImageIndex = Math.max(
    databaseData.images.findIndex((image) => image.isDefault),
    0,
  );
  const selectedImageIndex = databaseData.images.findIndex((image) => image.selected);
  const pendingDefaultImageIndex =
    selectedImageIndex >= 0 && selectedImageIndex !== defaultImageIndex
      ? selectedImageIndex
      : null;

  const handleImageSelection = (imageIndex: number) => {
    const selectedImage = databaseData.images[imageIndex];
    if (!selectedImage) {
      return;
    }
    setColor(selectedImage.spineColor);
    setDatabaseData((prev) => ({
      ...prev,
      spineColor: selectedImage.spineColor,
      images: prev.images.map((image, index) => ({
        ...image,
        selected: index === imageIndex,
      })),
    }));
  };

  const handleSetAsDefault = () => {
    if (pendingDefaultImageIndex == null) {
      return;
    }
    setDatabaseData((prev) => ({
      ...prev,
      spineColor: prev.images[pendingDefaultImageIndex]?.spineColor ?? prev.spineColor,
      images: prev.images.map((image, index) => ({
        ...image,
        isDefault: index === pendingDefaultImageIndex,
        selected: index === pendingDefaultImageIndex,
      })),
    }));
  };

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
      const spineColor = sRGBHex;
      setColor(spineColor);
      setDatabaseData((prev) => {
        const imageIndexToUpdate =
          prev.images.findIndex((image) => image.selected) >= 0
            ? prev.images.findIndex((image) => image.selected)
            : defaultImageIndex;
        return {
          ...prev,
          spineColor: spineColor,
          images: prev.images.map((image, index) =>
            index === imageIndexToUpdate ? { ...image, spineColor } : image,
          ),
        };
      });
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

  const { mutateAsync: databaseEdit } = trpc.database.edit.useMutation();
  const { mutateAsync: uploadCoverImage } =
    trpc.collect.uploadCoverImage.useMutation();
  const { mutateAsync: linkGenres } = trpc.genres.link.useMutation();
  const { mutateAsync: unlinkGenres } = trpc.genres.unlink.useMutation();
  const utils = trpc.useUtils();

  const convertFileToBase64 = (file: File): Promise<string> =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result;
        if (typeof result !== 'string') {
          reject(new Error('Unable to read uploaded file.'));
          return;
        }
        const [, dataBase64 = ''] = result.split(',');
        resolve(dataBase64);
      };
      reader.onerror = () => {
        reject(new Error('Unable to read uploaded file.'));
      };
      reader.readAsDataURL(file);
    });

  const handleUploadImage = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    setIsUploading(true);
    try {
      const dataBase64 = await convertFileToBase64(file);
      const uploadedImage = await uploadCoverImage({
        blockID: id,
        sortOrder: databaseData.images.length,
        fileName: file.name,
        mimeType: file.type,
        dataBase64,
      });
      setDatabaseData((prev) => ({
        ...prev,
        spineColor: color,
        images: [
          ...prev.images.map((image) => ({ ...image, selected: false })),
          {
            url: uploadedImage.url,
            selected: true,
            isDefault: false,
            spineColor: color,
          },
        ],
      }));
    } catch (error) {
      console.error('Failed to upload image for database edit', error);
    } finally {
      setIsUploading(false);
      event.target.value = '';
    }
  };

  const handleEditSubmit = async () => {
    const res = await databaseEdit({ type, item: databaseData });
    if ('error' in res) {
      setEdit(false);
      return;
    }

    const newGenres = databaseGenres;
    const linkGenresList: string[] = [];
    const unlinkGenresList: string[] = [];

    for (const genre of newGenres) {
      if (!initialGenres.includes(genre)) linkGenresList.push(genre);
    }
    if (linkGenresList.length > 0) {
      try {
        await linkGenres({ bookID: id, genres: linkGenresList });
        console.log('Linked genres');
      } catch {
        console.log('Genre link error');
      }
    }

    for (const genre of initialGenres) {
      if (!newGenres.includes(genre)) unlinkGenresList.push(genre);
    }
    if (unlinkGenresList.length > 0) {
      try {
        await unlinkGenres({ bookID: id, genres: unlinkGenresList });
      } catch {
        console.log('Genre unlink error');
      }
    }

    // Ensure per-item genre display updates by invalidating its query cache
    await utils.genres.getForBook.invalidate({ bookID: id });

    await handleGetMedia();
    setEdit(false);
  };

  return (
    <div className="border-darkpink bg-lightpink fixed top-1/2 left-1/2 z-100 flex -translate-x-1/2 -translate-y-1/2 flex-col content-center items-center justify-center gap-1 rounded-md border-3 p-2 text-2xl tracking-wider text-black">
      <h1>Editing: {titleRearrange(title)}</h1>

      <div
        className={`relative flex h-fit w-fit min-w-lg flex-col items-center gap-2.5 rounded-lg text-lg ${blockClasses[type]} mb-1`}
      >
        <div className="absolute top-1 left-1">{icons[type]}</div>
        <input
          ref={fileInputRef}
          type="file"
          accept="image/*"
          className="hidden"
          onChange={handleUploadImage}
          aria-label="Upload database image"
        />
        <MediaImageStrip
          mediaType={type}
          images={databaseData.images}
          className="m-2.5 mb-0 flex flex-row items-center gap-7.5"
          albumTileClassName="relative z-10 h-31 w-21 overflow-hidden rounded-sm"
          defaultTileClassName="relative z-10 h-31 w-21 overflow-hidden rounded-sm"
          showSelectionOverlay
          onImageClick={handleImageSelection}
          showUploadButton={type === 'book'}
          isUploading={isUploading}
          onUploadButtonClick={() => fileInputRef.current?.click()}
          uploadButtonLabel="Add uploaded database image"
          uploadSlotLabel="Uploading..."
        />
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
              name="pubYear"
              label="Pub Year"
              type={type}
              value={pubYear || ''}
              setDatabaseData={setDatabaseData}
            />
            <MyTextArea
              name="pageCount"
              label="Page Count"
              type={type}
              value={pageCount || ''}
              setDatabaseData={setDatabaseData}
            />
          </>
        ) : (
          <></>
        )}
        {type === 'book' ? (
          <GenreCheckboxes
            allGenres={genres}
            blockGenres={databaseGenres}
            handleGenreClick={handleGenreClick}
          />
        ) : (
          <></>
        )}
      </div>
      <div className="flex gap-2">
        <Button
          variant="primary"
          label="Close"
          onClick={() => setEdit(false)}
          width={100}
          fontSize={25}
        />
        <Button
          variant="primary"
          label="Submit Changes"
          onClick={handleEditSubmit}
          width={150}
          fontSize={25}
        />
        {pendingDefaultImageIndex != null ? (
          <Button
            variant="primary"
            label="Set as Default"
            onClick={handleSetAsDefault}
            width={165}
            fontSize={25}
          />
        ) : null}
      </div>
    </div>
  );
});

export default EditDatabaseBlock;
