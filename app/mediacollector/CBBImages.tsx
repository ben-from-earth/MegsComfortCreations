// interfaces and types
import Image from 'next/image';
import { CollectorFormData } from './collector-form/collectorFormSchema';
import { ChangeEvent, useRef, useState } from 'react';
import AddToPhotosIcon from '@mui/icons-material/AddToPhotos';

import { useFormContext } from 'react-hook-form';
import { trpc } from 'lib/trpc/client';

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
  const [brokenImageUrls, setBrokenImageUrls] = useState<
    Record<string, boolean>
  >({});
  const collectedData = watch('collectedData');
  const block = collectedData[blockID];
  if (!block) {
    return null;
  }
  const { type, images, isDatabase } = block;
  //setup connection to redux slice

  //add the image url to the database data (in the state) or removes it if its there already
  const handleClick = (
    image: { url: string; selected: boolean },
    imageIdx: number,
  ) => {
    if (!image.selected) {
      const newBlockImages = block.images.map((img, idx) => {
        if (idx === imageIdx) {
          return { ...img, selected: true };
        }
        return img;
      });
      setValue(`collectedData.${blockID}`, {
        ...block,
        images: newBlockImages,
      });
    } else {
      const newBlockImages = block.images.map((img, idx) => {
        if (idx === imageIdx) {
          return { ...img, selected: false };
        }
        return img;
      });
      setValue(`collectedData.${blockID}`, {
        ...block,
        images: newBlockImages,
      });
    }
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
      setValue(`collectedData.${blockID}.blockInfo.spineColor`, spineColor);
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
      const { url } = await uploadCoverImage({
        blockID: block.blockID,
        sortOrder: images.length,
        fileName: file.name,
        mimeType: file.type,
        dataBase64,
      });

      setValue(`collectedData.${blockID}`, {
        ...block,
        images: [...images, { url, selected: true }],
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
      {images.map((image, idx) => (
        <div
          className={`relative z-10 overflow-hidden rounded-sm ${
            type === 'album' ? 'h-31 w-31' : 'h-31 w-21'
          }`}
          key={image.url}
          onClick={() => {
            if (!isDatabase) {
              handleClick(image, idx);
            }
          }}
        >
          {brokenImageUrls[image.url] ? (
            <div className="absolute inset-0 flex items-center justify-center rounded-sm bg-red-600/65 p-2 text-center text-sm font-bold text-white">
              Image path broken
            </div>
          ) : (
            <Image
              className={
                type === 'album'
                  ? 'cursor-pointer object-cover outline-2'
                  : 'cursor-pointer'
              }
              src={image.url}
              alt={`${type} image`}
              fill
              sizes="(max-width: 640px) 33vw, (max-width: 1024px) 20vw, 200px"
              unoptimized
              loader={({ src }) => src}
              onError={() =>
                setBrokenImageUrls((previous) => ({
                  ...previous,
                  [image.url]: true,
                }))
              }
            />
          )}

          <div
            className={`pointer-events-none absolute inset-0 flex content-center items-center ${
              image.selected ? 'opacity-100' : 'opacity-0'
            }`}
          >
            <p className='-translate-x-1 -rotate-65 font-["Just_Another_Hand"] text-5xl font-bold tracking-wider text-[rgb(0,77,0)]'>
              Selected
            </p>
          </div>
        </div>
      ))}
      {showUploadSlot ? (
        <>
          <input
            ref={fileInputRef}
            type="file"
            accept="image/*"
            className="hidden"
            onChange={handleUploadImage}
            aria-label="Upload book image"
          />
          <button
            type="button"
            className="flex h-31 w-21 cursor-pointer items-center justify-center rounded-sm border-2 border-dashed border-black/60 bg-white/45 transition-opacity hover:opacity-85 disabled:cursor-not-allowed disabled:opacity-50"
            onClick={() => fileInputRef.current?.click()}
            disabled={isUploading}
            aria-label="Add uploaded book image"
          >
            {isUploading ? (
              <span className="text-center text-sm font-bold">Uploading...</span>
            ) : (
              <AddToPhotosIcon sx={{ fontSize: '2.25rem' }} />
            )}
          </button>
        </>
      ) : null}
    </div>
  );
}
