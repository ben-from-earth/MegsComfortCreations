// react, redux imports
import { Dispatch, SetStateAction } from 'react';

// necessary imports from collector state slice
import { mediaTypeDefinitions } from 'lib/state/slices/collectorSlice';

// interfaces and types
import { MediaType } from 'lib/interfaces/globalInterfaces';
import titleCollectionListConversion, {
  titleOutputObj,
} from 'lib/helpers/titleCollectionListConversion';

// component imports
import TextInput from '../components/TextInput';
import { useFormContext } from 'react-hook-form';
import { CollectorFormData } from '@//mediacollector/collector-form/collectorFormSchema';

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
  // mediaTypes can be moved to local state of page.tsx, because were updating the form state
  // only if media type box exists

  const { setValue } = useFormContext<CollectorFormData>();

  return (
    <form
      id="MediaInputForm"
      className="MediaInputs flex flex-col items-center gap-4 p-5 sm:grid sm:grid-cols-2"
    >
      {mediaTypes
        .filter((mediaType) => mediaType.show)
        .map(({ type, label }) => (
          <TextInput
            variant="multiline"
            key={type}
            label={`${label} Titles`}
            rows={5}
            onChange={(e) => {
              const titleSearchList = titleCollectionListConversion(
                e.target.value,
              );
              // for form change we can just do setValues('collectionData[type]', titleSearchList)
              // this removes need for setSearchData function
              setValue(`collectionData`, titleSearchList);
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
