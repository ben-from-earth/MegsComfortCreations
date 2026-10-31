import { ChangeEvent, useRef, useState } from 'react';
import { useFormContext, useWatch } from 'react-hook-form';
import { trpc } from 'lib/trpc/client';
import { MediaType } from 'lib/constants/media-types';
import { CollectorFormData } from './collector-form/collector-form-schema';
import {
  CoverImage,
  CoverTile,
  UploadTile,
  coverAspectClassName,
} from './cover-tile';
import { defaultCoverRibbonVariant } from '@/components/shared/default-cover-ribbon';
import CoversDialog from './covers-dialog';

const MAX_TILES_WITHOUT_OVERFLOW = 4;

export interface CBBImageProps {
  index: number;
  blockID: string;
  type: MediaType;
  isDatabase: boolean;
  isLocked: boolean;
  originalDefaultUrl?: string;
  title: string;
}

export default function CBBImages({
  index,
  blockID,
  type,
  isDatabase,
  isLocked,
  originalDefaultUrl,
  title,
}: CBBImageProps) {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [isDialogOpen, setIsDialogOpen] = useState(false);

  const { control, getValues, setValue } = useFormContext<CollectorFormData>();
  const { mutateAsync: uploadCoverImage } =
    trpc.collect.uploadCoverImage.useMutation();
  const images =
    useWatch({ control, name: `collectedData.${index}.images` }) ?? [];

  const handleImageSelection = (imageIdx: number) => {
    const currentImages = getValues(`collectedData.${index}.images`);
    const selectedImage = currentImages[imageIdx];
    if (!selectedImage) {
      return;
    }
    setValue(
      `collectedData.${index}.images`,
      currentImages.map((image, imageIndex) => ({
        ...image,
        selected: imageIndex === imageIdx,
      })),
    );
    setValue(
      `collectedData.${index}.blockInfo.spineColor`,
      selectedImage.spineColor,
    );
  };

  const handleSetDefault = (imageIdx: number) => {
    setValue(
      `collectedData.${index}.images`,
      getValues(`collectedData.${index}.images`).map((image, imageIndex) => ({
        ...image,
        isDefault: imageIndex === imageIdx,
      })),
    );
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
      const currentImages = getValues(`collectedData.${index}.images`);
      const spineColor = getValues(
        `collectedData.${index}.blockInfo.spineColor`,
      );
      const dataBase64 = await convertFileToBase64(file);
      const uploadedImage = await uploadCoverImage({
        blockID,
        sortOrder: currentImages.length,
        fileName: file.name,
        mimeType: file.type,
        dataBase64,
      });

      const uploadedSpineColor = isDatabase
        ? spineColor
        : (uploadedImage.palette[0] ?? spineColor);
      setValue(`collectedData.${index}.images`, [
        ...currentImages.map((image) => ({ ...image, selected: false })),
        {
          url: uploadedImage.url,
          selected: true,
          isDefault: false,
          spineColor: uploadedSpineColor,
          uploaded: true,
          palette: uploadedImage.palette,
        },
      ]);
      setValue(
        `collectedData.${index}.blockInfo.spineColor`,
        uploadedSpineColor,
      );
    } catch (error) {
      console.error('Failed to upload custom book image', error);
    } finally {
      setIsUploading(false);
      event.target.value = '';
    }
  };

  const showUploadSlot = type === 'book';
  const needsCover = !isDatabase && !images.some((image) => image.selected);

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
          needsCover
            ? 'outline-[1.5px] outline-offset-[6px] outline-[#b0533a] outline-dashed'
            : ''
        }`}
      >
        {visibleImageIndexes.map((imageIndex) => {
          const image = images[imageIndex];
          return (
            <CoverTile
              key={`${image.url}-${imageIndex}`}
              url={image.url}
              mediaType={type}
              coverNumber={imageIndex + 1}
              isSelected={image.selected}
              defaultRibbon={
                isDatabase
                  ? defaultCoverRibbonVariant(image, originalDefaultUrl)
                  : null
              }
              onClick={() => handleImageSelection(imageIndex)}
              onSetDefault={
                isDatabase && !isLocked && image.selected && !image.isDefault
                  ? () => handleSetDefault(imageIndex)
                  : undefined
              }
            />
          );
        })}
        {hasOverflow ? (
          <button
            type="button"
            onClick={() => setIsDialogOpen(true)}
            className={`relative flex w-full cursor-pointer items-center justify-center overflow-hidden rounded-md border-0 p-0 ${coverAspectClassName(type)}`}
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
            onClick={() => fileInputRef.current?.click()}
          />
        ) : null}
      </div>
      {isDialogOpen ? (
        <CoversDialog
          title={title}
          isDatabase={isDatabase}
          mediaType={type}
          images={images}
          isUploading={isUploading}
          originalDefaultUrl={originalDefaultUrl}
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
