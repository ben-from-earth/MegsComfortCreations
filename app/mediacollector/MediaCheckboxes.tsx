// react, redux imports
import { Dispatch, SetStateAction } from 'react';
import { useDispatch } from 'react-redux';

// necessary imports from collector state slice
import {
  mediaTypeDefinitions,
  setChecks,
} from 'lib/state/slices/collectorSlice';

// interfaces and types
import { MediaType } from 'lib/interfaces/globalInterfaces';
import { titleOutputObj } from 'lib/helpers/titleCollectionListConversion';

interface MediaCheckboxesProps {
  mediaTypes: mediaTypeDefinitions[];
  setSearchData: Dispatch<
    SetStateAction<
      {
        type: MediaType;
        titleSearchList: titleOutputObj[];
      }[]
    >
  >;
}

export default function MediaCheckboxes({
  mediaTypes,
  setSearchData,
}: MediaCheckboxesProps) {
  //setup connection to redux slice
  const dispatch = useDispatch();

  return (
    <div className="m-6 flex flex-row content-center gap-5">
      {mediaTypes.map(({ type, label, show }, idx) => (
        <label key={type} className="text-3xl tracking-wider">
          <input
            checked={show}
            className="m-1.5"
            id={`${idx}`}
            type="checkbox"
            onChange={() => {
              dispatch(setChecks(idx));
              setSearchData((prev) => {
                return prev.map((item, i) =>
                  i === idx
                    ? { ...prev[i], type: item.type, titleSearchList: [] }
                    : prev[i],
                );
              });
            }}
          />
          {`${label}s`}
        </label>
      ))}
    </div>
  );
}
