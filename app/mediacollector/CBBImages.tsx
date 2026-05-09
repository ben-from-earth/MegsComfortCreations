// interfaces and types
import { CollectorFormData } from './collector-form/collectorFormSchema';
import { ChangeEvent, useRef, useState } from 'react';

import { useFormContext } from 'react-hook-form';
import { trpc } from 'lib/trpc/client';
import MediaImageStrip from '@/shared/MediaImageStrip';

export interface CBBImageProps {
  blockID: number;
  spineColor: string;
}

export default function CBBImages({ blockID, spineColor }: CBBImageProps) {
  //set local state for spine color
  const [color, setColor] = useState(spineColor);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [hasUploadedCustomImage, setHasUploadedCustomImage] = useState(false);

  const { watch, setValue } = useFormContext<CollectorFormData>();
  const { mutateAsync: uploadCoverImage } =
    trpc.collect.uploadCoverImage.useMutation();
  const collectedData = watch('collectedData');
  const block = collectedData[blockID];
  if (!block) {
    return null;
  }
  const { type, images, isDatabase } = block;
  //setup connection to redux slice

  const handleImageSelection = (imageIdx: number) => {
    const selectedImage = block.images[imageIdx];
    if (!selectedImage) {
      return;
    }
    const nextImages = block.images.map((image, index) => ({
      ...image,
      selected: index === imageIdx,
    }));
    setColor(selectedImage.spineColor);
    setValue(`collectedData.${blockID}`, {
      ...block,
      images: nextImages,
      blockInfo: {
        ...block.blockInfo,
        spineColor: selectedImage.spineColor,
      },
    });
  };

  const handleColorPick = async (blockID: number) => {
    if (!window.EyeDropper) {
      console.log('EyeDropper API not supported in this browser');
      return;
    }
    const eyeDropper = new window.EyeDropper();
    try {
      const { sRGBHex } = await eyeDropper.open();
      const spineColor = sRGBHex;
      setColor(spineColor);
      const selectedImageIndex = block.images.findIndex((image) => image.selected);
      const imageIndexToUpdate = selectedImageIndex >= 0 ? selectedImageIndex : 0;
      const nextImages = block.images.map((image, index) =>
        index === imageIndexToUpdate ? { ...image, spineColor } : image,
      );
      setValue(`collectedData.${blockID}`, {
        ...block,
        images: nextImages,
        blockInfo: {
          ...block.blockInfo,
          spineColor,
        },
      });
    } catch (e) {
      console.log(e);
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
      const dataBase64 = await convertFileToBase64(file);
      const uploadedImage = await uploadCoverImage({
        blockID: block.blockID,
        sortOrder: images.length,
        fileName: file.name,
        mimeType: file.type,
        dataBase64,
      });

      setValue(`collectedData.${blockID}`, {
        ...block,
        images: [
          ...images.map((image) => ({ ...image, selected: false })),
          {
            url: uploadedImage.url,
            selected: true,
            isDefault: images.length === 0,
            spineColor: color,
          },
        ],
        blockInfo: {
          ...block.blockInfo,
          spineColor: color,
        },
      });
      setHasUploadedCustomImage(true);
    } catch (error) {
      console.error('Failed to upload custom book image', error);
    } finally {
      setIsUploading(false);
      event.target.value = '';
    }
  };

  const showUploadSlot = type === 'book' && !isDatabase && !hasUploadedCustomImage;

  return (
    <div className="mx-10 mt-2.5 flex flex-row items-center gap-3">
      {type !== 'album' ? (
        <div
          className="h-full w-5 cursor-pointer rounded-sm"
          style={{ backgroundColor: color }}
          onClick={() => handleColorPick(blockID)}
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
