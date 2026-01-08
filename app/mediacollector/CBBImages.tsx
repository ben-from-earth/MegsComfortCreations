// react, redux imports
import { useAppDispatch } from 'lib/state/store';

//necessary imports from png collection state slice
import {
  addToPNGCollectionList,
  removeFromPNGCollectionList,
} from 'lib/state/slices/pngCollectionSlice';

// interfaces and types
import { MediaType } from 'lib/interfaces/globalInterfaces';
import Image from 'next/image';
import { CollectorFormData } from './collector-form/collectorFormSchema';

import { useFormContext } from 'react-hook-form';

export interface CBBImageProps {
  blockID: number;
}

export default function CBBImages({ blockID }: CBBImageProps) {
  const { watch, setValue } = useFormContext<CollectorFormData>();
  const collectedData = watch('collectedData');
  const block = collectedData[blockID];
  const dispatch = useAppDispatch();
  if (!block) {
    return null;
  }
  const {
    type,
    images,
    blockInfo: { spineColor },
    isDatabase,
  } = block;
  //setup connection to redux slice

  //add the image url to the database data (in the state) or removes it if its there already
  const handleClick = (
    image: { url: string; selected: boolean },
    type: MediaType,
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

      dispatch(addToPNGCollectionList({ url: image.url, type, spineColor }));
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
      dispatch(removeFromPNGCollectionList({ url: image.url }));
    }
  };

  return (
    <div className="mx-10 mt-2.5 flex flex-row items-center gap-5">
      {images.map((image, idx) => (
        <div
          className={`relative z-10 overflow-hidden rounded-sm ${
            type === 'album' ? 'w-31' : 'h-31 w-21'
          }`}
          key={image.url}
          onClick={() => {
            if (!isDatabase) {
              handleClick(image, type, idx);
            }
          }}
        >
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
          />

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
    </div>
  );
}
