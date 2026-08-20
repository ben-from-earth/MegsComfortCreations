import { ChangeEvent, useContext, useRef, useState } from 'react';
import { useFormContext } from 'react-hook-form';
import GenreContext from 'lib/context/GenreContext';
import GenreCheckboxes from '@/mediacollector/GenreCheckboxes';
import Button from '@/components/ui/Button';
import MediaImageStrip from '@/components/shared/MediaImageStrip';
import {
  Form,
  FormControl,
  FormField,
  FormItem,
  FormLabel,
  FormMessage,
} from '@/components/ui/form';
import { trpc } from 'lib/trpc/client';
import { titleRearrange } from 'lib/helpers/titleRearrange';
import { MediaType } from 'lib/constants/mediaTypes';
import { blockClasses, icons } from 'lib/constants/typeBlockStyles';
import type { MediaItemForm } from '@/mediacollector/collector-form/mediaItemFormSchema';
import { useMediaItemForm } from './use-media-item-form';

export interface MediaItemAddEditProps {
  item: MediaItemForm;
  onClose: () => void;
}

const NUMBER_FIELD_NAMES = ['pubYear', 'pageCount'] as const;

type MediaItemTextFieldName = 'title' | 'author' | 'pubYear' | 'pageCount';

function isNumberFieldName(
  name: MediaItemTextFieldName,
): name is (typeof NUMBER_FIELD_NAMES)[number] {
  return NUMBER_FIELD_NAMES.some((fieldName) => fieldName === name);
}

function MediaItemTextField({
  name,
  label,
  type,
}: {
  name: MediaItemTextFieldName;
  label: string;
  type: MediaType;
}) {
  const { control } = useFormContext<MediaItemForm>();
  const labelClass =
    type === 'book'
      ? 'w-25 content-center text-right text-2xl'
      : 'w-15 content-center text-right text-2xl';

  return (
    <FormField
      control={control}
      name={`blockInfo.${name}`}
      render={({ field }) => (
        <FormItem className="grid grid-cols-[max-content_1fr] gap-x-3 gap-y-1 p-2">
          <FormLabel className={labelClass}>
            {label}:
          </FormLabel>
          <FormControl>
            <textarea
              className="w-2xs content-center rounded-sm bg-white pl-2 text-black"
              name={field.name}
              value={field.value == null ? '' : String(field.value)}
              onBlur={field.onBlur}
              onChange={(event) => {
                if (isNumberFieldName(name)) {
                  const parsed = Number(event.target.value);
                  field.onChange(
                    event.target.value.trim() === '' || !Number.isFinite(parsed)
                      ? null
                      : parsed,
                  );
                  return;
                }
                field.onChange(event.target.value);
              }}
            />
          </FormControl>
          <div className="col-start-2">
            <FormMessage />
          </div>
        </FormItem>
      )}
    />
  );
}

function MediaItemAddEditFields({
  type,
  blockID,
  onClose,
  submitError,
}: {
  type: MediaType;
  blockID: string;
  onClose: () => void;
  submitError: string | null;
}) {
  const allGenres = useContext(GenreContext);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const { control, getValues, setValue, watch } =
    useFormContext<MediaItemForm>();
  const { mutateAsync: uploadCoverImage } =
    trpc.collect.uploadCoverImage.useMutation();

  const images = watch('images');
  const spineColor = watch('blockInfo.spineColor');
  const genres = watch('blockInfo.genres') ?? [];

  const defaultImageIndex = Math.max(
    images.findIndex((image) => image.isDefault),
    0,
  );
  const selectedImageIndex = images.findIndex((image) => image.selected);
  const pendingDefaultImageIndex =
    selectedImageIndex >= 0 && selectedImageIndex !== defaultImageIndex
      ? selectedImageIndex
      : null;

  const handleImageSelection = (imageIndex: number) => {
    const currentImages = getValues('images');
    const selectedImage = currentImages[imageIndex];
    if (!selectedImage) {
      return;
    }
    setValue(
      'images',
      currentImages.map((image, index) => ({
        ...image,
        selected: index === imageIndex,
      })),
    );
    setValue('blockInfo.spineColor', selectedImage.spineColor);
  };

  const handleSetAsDefault = () => {
    if (pendingDefaultImageIndex == null) {
      return;
    }
    const currentImages = getValues('images');
    const nextDefault = currentImages[pendingDefaultImageIndex];
    setValue(
      'images',
      currentImages.map((image, index) => ({
        ...image,
        isDefault: index === pendingDefaultImageIndex,
        selected: index === pendingDefaultImageIndex,
      })),
    );
    if (nextDefault) {
      setValue('blockInfo.spineColor', nextDefault.spineColor);
    }
  };

  const handleColorPick = async () => {
    if (!window.EyeDropper) {
      console.log('EyeDropper API not supported in this browser');
      return;
    }
    const eyeDropper = new window.EyeDropper();
    try {
      const { sRGBHex } = await eyeDropper.open();
      const currentImages = getValues('images');
      const selectedIndex = currentImages.findIndex((image) => image.selected);
      const imageIndexToUpdate = selectedIndex >= 0 ? selectedIndex : 0;
      setValue(
        'images',
        currentImages.map((image, index) =>
          index === imageIndexToUpdate
            ? { ...image, spineColor: sRGBHex }
            : image,
        ),
      );
      setValue('blockInfo.spineColor', sRGBHex);
    } catch (error) {
      console.log(error);
    }
  };

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
      const currentImages = getValues('images');
      const currentSpineColor = getValues('blockInfo.spineColor');
      const dataBase64 = await convertFileToBase64(file);
      const uploadedImage = await uploadCoverImage({
        blockID,
        sortOrder: currentImages.length,
        fileName: file.name,
        mimeType: file.type,
        dataBase64,
      });
      setValue('images', [
        ...currentImages.map((image) => ({ ...image, selected: false })),
        {
          url: uploadedImage.url,
          selected: true,
          isDefault: false,
          spineColor: currentSpineColor,
        },
      ]);
    } catch (error) {
      console.error('Failed to upload image for database edit', error);
    } finally {
      setIsUploading(false);
      event.target.value = '';
    }
  };

  const handleGenreClick = (genreText: string, checked: boolean) => {
    const currentGenres = getValues('blockInfo.genres') ?? [];
    setValue(
      'blockInfo.genres',
      checked
        ? [...currentGenres, genreText]
        : currentGenres.filter((genre) => genre !== genreText),
    );
  };

  return (
    <>
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
        <FormField
          control={control}
          name="images"
          render={() => (
            <FormItem className="flex flex-col items-center">
              <MediaImageStrip
                mediaType={type}
                images={images}
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
              <FormMessage />
            </FormItem>
          )}
        />
        {type !== 'album' ? (
          <div
            className="h-5 w-1/2 cursor-pointer"
            style={{ backgroundColor: spineColor }}
            onClick={() => handleColorPick()}
          ></div>
        ) : null}

        <MediaItemTextField name="title" label="Title" type={type} />
        {type === 'book' ? (
          <>
            <MediaItemTextField name="author" label="Author" type={type} />
            <MediaItemTextField name="pubYear" label="Pub Year" type={type} />
            <MediaItemTextField
              name="pageCount"
              label="Page Count"
              type={type}
            />
            <GenreCheckboxes
              allGenres={allGenres}
              blockGenres={genres}
              handleGenreClick={handleGenreClick}
            />
          </>
        ) : null}
      </div>
      {submitError ? (
        <p className="m-0 max-w-lg text-center font-['Just_Another_Hand'] text-2xl tracking-wider text-red-600">
          {submitError}
        </p>
      ) : null}
      <div className="flex gap-2">
        <Button
          variant="primary"
          label="Close"
          onClick={onClose}
          width={100}
          fontSize={25}
        />
        <Button
          type="submit"
          variant="primary"
          label="Submit Changes"
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
    </>
  );
}

export default function MediaItemAddEdit({
  item,
  onClose,
}: MediaItemAddEditProps) {
  const { form, formId, onSubmit, submitError } = useMediaItemForm({
    item,
    onClose,
  });

  return (
    <Form {...form}>
      <form
        id={formId}
        onSubmit={form.handleSubmit(onSubmit)}
        className="border-darkpink bg-lightpink fixed top-1/2 left-1/2 z-100 flex -translate-x-1/2 -translate-y-1/2 flex-col content-center items-center justify-center gap-1 rounded-md border-3 p-2 text-2xl tracking-wider text-black"
      >
        <h1>Editing: {titleRearrange(item.blockInfo.title)}</h1>
        <MediaItemAddEditFields
          type={item.type}
          blockID={item.blockID}
          onClose={onClose}
          submitError={submitError}
        />
      </form>
    </Form>
  );
}
