// react, redux imports
import { useState } from 'react';
import { useAppDispatch } from '@/lib/state/store';

// necessary imports from database data state slice
import {
  addToDatabaseData,
  removeFromDatabaseData,
} from '@/lib/state/slices/databaseDataSlice';

//necessary imports from png collection state slice
import {
  addToPNGCollectionList,
  removeFromPNGCollectionList,
} from '@/lib/state/slices/pngCollectionSlice';

// interfaces and types
import { MediaType } from '@/lib/interfaces/globalInterfaces';

export interface CBBImageProps {
  images: string[];
  isDatabase: boolean;
  blockID: string;
  type: MediaType;
  spineColor: string;
}

export default function CBBImages({
  images,
  isDatabase,
  blockID,
  type,
  spineColor,
}: CBBImageProps) {
  //setup connection to redux slice
  const dispatch = useAppDispatch();

  //set up a local state to an array with an index for each image slot (current: 3) and set to false
  //this is for click tracking and the "selected" style and setting the image as a block datapoint
  const [clicked, setClicked] = useState<boolean[]>(() =>
    Array(images.length).fill(false),
  );

  //add the image url to the database data (in the state) or removes it if its there already
  const handleClick = (
    blockID: string,
    type: MediaType,
    idx: number,
    src: string,
  ) => {
    const next = !clicked[idx];
    setClicked((prev) =>
      prev.map((b, itemIndex) => (itemIndex === idx ? next : b)),
    );
    if (next) {
      dispatch(
        addToDatabaseData({
          type,
          src,
          idx,
          blockID,
        }),
      );
      dispatch(addToPNGCollectionList({ url: src, type, spineColor }));
    } else {
      dispatch(removeFromDatabaseData({ blockID, type, idx }));
      dispatch(removeFromPNGCollectionList({ url: src }));
    }
  };

  return (
    <div className="mx-10 mt-2.5 flex flex-row items-center gap-5">
      {images.map((src, idx) => (
        <div
          className="relative z-10 overflow-hidden rounded-sm"
          key={src}
          onClick={() => {
            if (!isDatabase) {
              handleClick(blockID, type, idx, src);
            }
          }}
        >
          <img
            className={
              type === 'album'
                ? 'block w-31 cursor-pointer object-cover outline-2'
                : 'block h-31 w-21 cursor-pointer'
            }
            src={src}
          ></img>

          <div
            className={`pointer-events-none absolute inset-0 flex content-center items-center bg-[rgba(0,200,0,0.5)] ${
              clicked[idx] ? 'opacity-100' : 'opacity-0'
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
