import { ChangeEvent, useRef, useState } from 'react';
import CloseIcon from '@mui/icons-material/Close';
import { useFormContext, useFormState, useWatch } from 'react-hook-form';
import { trpc } from 'lib/trpc/client';
import { readFileAsBase64 } from 'lib/helpers/read-file-as-base64';
import { MediaType } from 'lib/constants/media-types';
import { defaultCoverRibbonVariant } from '@/components/shared/default-cover-ribbon';
import {
  CoverImage,
  CoverTile,
  UploadTile,
  collectorCoverAspectClassName,
} from '@/mediacollector/cover-tile';
import CoversDialog from '@/mediacollector/covers-dialog';
import type { MediaItemForm } from '@/mediacollector/collector-form/media-item-form-schema';
import { removeLibraryCover } from './cover-edit-rules';

const MAX_TILES_WITHOUT_OVERFLOW = 4;

export default function LibraryItemCovers({
  type,
  blockID,
  title,
  isLocked,
  originalDefaultUrl,
  isCoverListChanged,
}: {
  type: MediaType;
  blockID: string;
  title: string;
  isLocked: boolean;
  originalDefaultUrl: string | undefined;
  isCoverListChanged: boolean;
}) {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const { control, getValues, setValue } = useFormContext<MediaItemForm>();
  const { mutateAsync: uploadCoverImage } =
    trpc.collect.uploadCoverImage.useMutation();
  const images = useWatch({ control, name: 'images' }) ?? [];
  const { errors } = useFormState({ control, name: 'images' });
  const imagesError = errors.images?.message;

  const applyImages = (nextImages: MediaItemForm['images']) => {
    setValue('images', nextImages);
    const selectedImage = nextImages.find((image) => image.selected);
    if (selectedImage) {
      setValue('blockInfo.spineColor', selectedImage.spineColor);
    }
  };

  const handleImageSelection = (imageIndex: number) => {
    applyImages(
      getValues('images').map((image, index) => ({
        ...image,
        selected: index === imageIndex,
      })),
    );
  };

  const handleSetDefault = (imageIndex: number) => {
    setValue(
      'images',
      getValues('images').map((image, index) => ({
        ...image,
        isDefault: index === imageIndex,
      })),
    );
  };

  const handleUploadImage = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    setIsUploading(true);
    try {
      const currentImages = getValues('images');
      const spineColor = getValues('blockInfo.spineColor');
      const dataBase64 = await readFileAsBase64(file);
      const uploadedImage = await uploadCoverImage({
        blockID,
        sortOrder: currentImages.length,
        fileName: file.name,
        mimeType: file.type,
        dataBase64,
      });
      applyImages([
        ...currentImages.map((image) => ({ ...image, selected: false })),
        {
          url: uploadedImage.url,
          selected: true,
          isDefault: false,
          spineColor,
          codeColor: null,
          uploaded: true,
          palette: uploadedImage.palette,
        },
      ]);
    } catch (error) {
      console.error('Failed to upload cover for library item', error);
    } finally {
      setIsUploading(false);
      event.target.value = '';
    }
  };

  const showUploadSlot =
    (type === 'book' || type === 'movie' || type === 'album') && !isLocked;
  const tileAspectClassName = collectorCoverAspectClassName(type);
  const hasOverflow = images.length > MAX_TILES_WITHOUT_OVERFLOW;
  const selectedImageIndex = images.findIndex((image) => image.selected);
  const visibleImageIndexes = hasOverflow
    ? selectedImageIndex > 2
      ? [0, 1, selectedImageIndex]
      : [0, 1, 2]
    : images.map((_, imageIndex) => imageIndex);
  const overflowPreviewImage = hasOverflow
    ? images.find((_, imageIndex) => !visibleImageIndexes.includes(imageIndex))
    : undefined;

  return (
    <>
      <input
        ref={fileInputRef}
        type="file"
        accept="image/*"
        className="hidden"
        onChange={handleUploadImage}
      />
      <div
        role="radiogroup"
        className={`mx-4.5 mb-4 grid grid-cols-5 gap-2.5 rounded-lg ${
          isCoverListChanged
            ? 'outline-[1.5px] outline-offset-[6px] outline-[#c9973f]'
            : ''
        }`}
      >
        {visibleImageIndexes.map((imageIndex) => {
          const image = images[imageIndex];
          const canDelete = !isLocked && image.selected && images.length > 1;
          return (
            <div key={`${image.url}-${imageIndex}`} className="relative">
              <CoverTile
                url={image.url}
                mediaType={type}
                coverNumber={imageIndex + 1}
                isSelected={image.selected}
                defaultRibbon={defaultCoverRibbonVariant(
                  image,
                  originalDefaultUrl,
                )}
                aspectClassName={tileAspectClassName}
                codeColor={image.codeColor}
                onClick={() => handleImageSelection(imageIndex)}
                onSetDefault={
                  !isLocked && image.selected && !image.isDefault
                    ? () => handleSetDefault(imageIndex)
                    : undefined
                }
              />
              {canDelete ? (
                <button
                  type="button"
                  aria-label={`Delete cover ${imageIndex + 1}`}
                  title="Delete cover"
                  onClick={() =>
                    applyImages(
                      removeLibraryCover(getValues('images'), imageIndex),
                    )
                  }
                  className="absolute right-1.25 bottom-1.25 z-20 flex h-5.5 w-5.5 cursor-pointer items-center justify-center rounded-full border-0 bg-white p-0 text-[#9a3b26] shadow-[0_1px_3px_rgba(0,0,0,0.25)]"
                >
                  <CloseIcon sx={{ fontSize: 14 }} />
                </button>
              ) : null}
            </div>
          );
        })}
        {hasOverflow ? (
          <button
            type="button"
            onClick={() => setIsDialogOpen(true)}
            className={`relative flex w-full cursor-pointer items-center justify-center overflow-hidden rounded-md border-0 p-0 ${tileAspectClassName}`}
          >
            {overflowPreviewImage ? (
              <CoverImage
                url={overflowPreviewImage.url}
                alt={`${type} cover`}
              />
            ) : null}
            <span className="absolute inset-0 bg-[rgba(43,40,38,0.66)]" />
            <span className="relative flex flex-col items-center gap-0.5 text-white">
              <span className="text-lg font-extrabold">
                +{images.length - visibleImageIndexes.length}
              </span>
              <span className="text-[11px] font-extrabold">
                See all {images.length}
              </span>
            </span>
          </button>
        ) : null}
        {showUploadSlot ? (
          <UploadTile
            mediaType={type}
            isUploading={isUploading}
            isDisabled={isLocked}
            aspectClassName={tileAspectClassName}
            onClick={() => fileInputRef.current?.click()}
          />
        ) : null}
      </div>
      {imagesError ? (
        <p className="mx-4.5 mt-0 mb-3 text-xs font-bold text-red-600">
          {imagesError}
        </p>
      ) : null}
      {isDialogOpen ? (
        <CoversDialog
          title={title}
          isDatabase
          mediaType={type}
          images={images}
          isUploading={isUploading}
          originalDefaultUrl={originalDefaultUrl}
          aspectClassName={tileAspectClassName}
          onCancel={() => setIsDialogOpen(false)}
          onConfirm={(imageIndex) => {
            handleImageSelection(imageIndex);
            setIsDialogOpen(false);
          }}
          onUploadClick={
            showUploadSlot && !isLocked
              ? () => {
                  setIsDialogOpen(false);
                  fileInputRef.current?.click();
                }
              : undefined
          }
        />
      ) : null}
    </>
  );
}
