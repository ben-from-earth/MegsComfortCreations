// react, redux imports
import { Dispatch, SetStateAction } from 'react';

// library imports
import { TextField } from '@mui/material';

// necessary imports from collector state slice
import { mediaTypeDefinitions } from '@/lib/state/slices/collectorSlice';

// interfaces and types
import { MediaType } from '@/lib/interfaces/globalInterfaces';
import titleCollectionListConversion, {
  titleOutputObj,
} from '@/lib/helpers/titleCollectionListConversion';

interface MediaInputProps {
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

export default function MediaInputs({
  mediaTypes,
  setSearchData,
}: MediaInputProps) {
  return (
    <form
      id="MediaInputForm"
      className="MediaInputs flex flex-col items-center gap-4 p-5 sm:grid sm:grid-cols-2"
    >
      {mediaTypes
        .filter((mediaType) => mediaType.show)
        .map(({ type, label }) => (
          <TextField
            className="w-90 rounded-sm bg-white"
            id={`outlined-multiline-static ${label}`}
            multiline
            key={type}
            label={`${label} Titles`}
            slotProps={{
              inputLabel: {
                sx: {
                  '&.MuiInputLabel-shrink': {
                    backgroundColor: 'white',
                    borderRadius: '8px',
                    px: '10px',
                    color: 'rgb(0,0,0, 0.5)',
                    transform: 'translate(6px, -8px) scale(0.75)',
                  },
                },
              },
            }}
            rows={5}
            onChange={(e) => {
              const titleSearchList = titleCollectionListConversion(
                e.target.value,
              );
              setSearchData((prev) => {
                return prev.map((media) =>
                  media.type === type ? { type: type, titleSearchList } : media,
                );
              });
            }}
          />
        ))}
    </form>
  );
}
