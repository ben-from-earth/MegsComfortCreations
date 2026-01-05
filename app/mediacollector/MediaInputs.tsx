// necessary imports from collector state slice
import { mediaTypeDefinitions } from 'lib/state/slices/collectorSlice';

// interfaces and types
import titleCollectionListConversion from 'lib/helpers/titleCollectionListConversion';

// component imports
import TextInput from '@/shared/TextInput';
import { useFormContext } from 'react-hook-form';
import { CollectorFormData } from '@//mediacollector/collector-form/collectorFormSchema';

interface MediaInputProps {
  mediaTypes: mediaTypeDefinitions[];
}

export default function MediaInputs({ mediaTypes }: MediaInputProps) {
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
              setValue(`collectionList.${type}`, titleSearchList);
            }}
          />
        ))}
    </form>
  );
}
