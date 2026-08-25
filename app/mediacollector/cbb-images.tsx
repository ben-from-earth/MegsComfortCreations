import { ChangeEvent, useRef, useState } from 'react';
import { useFormContext } from 'react-hook-form';
import { trpc } from 'lib/trpc/client';
import MediaImageStrip from '@/components/shared/media-image-strip';
import { CollectorFormData } from './collector-form/collector-form-schema';
import { MediaType } from 'lib/constants/media-types';

export interface CBBImageProps {
  index: number;
  blockID: string;
  type: MediaType;
  isDatabase: boolean;
  spineColor: string;
}

export default function CBBImages({
  index,
  blockID,
  type,
  isDatabase,
  spineColor,
}: CBBImageProps) {
  const [color, setColor] = useState(spineColor);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [hasUploadedCustomImage, setHasUploadedCustomImage] = useState(false);

  const { getValues, setValue, watch } = useFormContext<CollectorFormData>();
  const { mutateAsync: uploadCoverImage } =
    trpc.collect.uploadCoverImage.useMutation();
  const images = watch(`collectedData.${index}.images`);

  if (!images) {
    return null;
  }

  const handleImageSelection = (imageIdx: number) => {
    const currentImages = getValues(`collectedData.${index}.images`);
    const selectedImage = currentImages[imageIdx];
    if (!selectedImage) {
      return;
    }
    setColor(selectedImage.spineColor);
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

  const handleColorPick = async () => {
    if (!window.EyeDropper) {
      console.log('EyeDropper API not supported in this browser');
      return;
    }
    const eyeDropper = new window.EyeDropper();
    try {
      const { sRGBHex } = await eyeDropper.open();
      const nextSpineColor = sRGBHex;
      setColor(nextSpineColor);
      const currentImages = getValues(`collectedData.${index}.images`);
      const selectedImageIndex = currentImages.findIndex(
        (image) => image.selected,
      );
      const imageIndexToUpdate =
        selectedImageIndex >= 0 ? selectedImageIndex : 0;
      setValue(
        `collectedData.${index}.images`,
        currentImages.map((image, imageIndex) =>
          imageIndex === imageIndexToUpdate
            ? { ...image, spineColor: nextSpineColor }
            : image,
        ),
      );
      setValue(`collectedData.${index}.blockInfo.spineColor`, nextSpineColor);
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
      const currentImages = getValues(`collectedData.${index}.images`);
      const dataBase64 = await convertFileToBase64(file);
      const uploadedImage = await uploadCoverImage({
        blockID,
        sortOrder: currentImages.length,
        fileName: file.name,
        mimeType: file.type,
        dataBase64,
      });

      setValue(`collectedData.${index}.images`, [
        ...currentImages.map((image) => ({ ...image, selected: false })),
        {
          url: uploadedImage.url,
          selected: true,
          isDefault: currentImages.length === 0,
          spineColor: color,
        },
      ]);
      setValue(`collectedData.${index}.blockInfo.spineColor`, color);
      setHasUploadedCustomImage(true);
    } catch (error) {
      console.error('Failed to upload custom book image', error);
    } finally {
      setIsUploading(false);
      event.target.value = '';
    }
  };

  const showUploadSlot =
    type === 'book' && !isDatabase && !hasUploadedCustomImage;

  return (
    <div className="mx-10 mt-2.5 flex flex-row items-center gap-3">
      {type !== 'album' ? (
        <div
          className="h-full w-5 cursor-pointer rounded-sm"
          style={{ backgroundColor: color }}
          onClick={() => handleColorPick()}
        ></div>
      ) : null}
      <input
        ref={fileInputRef}
        type="file"
        accept="image/*"
        className="hidden"
        onChange={handleUploadImage}
        aria-label="Upload book image"
      />
      <MediaImageStrip
        mediaType={type}
        images={images}
        showSelectionOverlay
        onImageClick={(imageIndex) => {
          handleImageSelection(imageIndex);
        }}
        showUploadButton={showUploadSlot}
        isUploading={isUploading}
        onUploadButtonClick={() => fileInputRef.current?.click()}
        uploadButtonLabel="Add uploaded book image"
        uploadSlotLabel="Uploading..."
      />
    </div>
  );
}
