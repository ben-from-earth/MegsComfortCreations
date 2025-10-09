import { useState } from 'react';
import { useDispatch } from 'react-redux';
import {
  addToDatabaseData,
  removeFromDatabaseData,
} from '@/state/databaseDataSlice';
import {
  addToPNGCollectionList,
  removeFromPNGCollectionList,
} from '@/state/pngCollectionSlice';

const CCBImages = ({ images, isDatabase, blockID, type, color }) => {
  //setup connection to redux slice
  const dispatch = useDispatch();
  //set up a local state to an array with an index for each image slot (current: 3) and set to false
  //this is for click tracking and the "selected" style and setting the image as a block datapoint
  const [clicked, setClicked] = useState(() =>
    Array(images.length).fill(false),
  );

  // function for handling an image click
  //this adds the image url to the database data (in the state) or removes it if its there already
  const handleClick = (blockID, type, idx, src) => {
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
      dispatch(addToPNGCollectionList({ url: src, type, spine_color: color }));
    } else {
      dispatch(removeFromDatabaseData({ blockID, type, idx }));
      dispatch(removeFromPNGCollectionList({ url: src }));
    }
  };
  return (
    <div className="gap-7.5 m-2.5 mb-0 flex flex-row items-center">
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
                ? 'w-31 block cursor-pointer object-cover outline-2'
                : 'w-21 h-31 block cursor-pointer'
            }
            src={src}
          ></img>
          <div
            className={`pointer-events-none absolute inset-0 flex content-center items-center bg-[rgba(0,200,0,0.5)] ${clicked[idx] ? 'opacity-100' : 'opacity-0'}`}
          >
            <p className='-rotate-65 -translate-x-1 font-["Just_Another_Hand"] text-5xl font-bold tracking-wider text-[rgb(0,77,0)]'>
              Selected
            </p>
          </div>
        </div>
      ))}
    </div>
  );
};

export default CCBImages;
