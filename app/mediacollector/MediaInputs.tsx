// interfaces and types
import titleCollectionListConversion from 'lib/helpers/titleCollectionListConversion';

// component imports
import TextInput from '@/shared/TextInput';
import { useFormContext } from 'react-hook-form';
import { CollectorFormData } from '@/mediacollector/collector-form/collectorFormSchema';

// interface MediaInputProps {
//   mediaTypes: mediaTypeDefinitions[];
// }

export default function MediaInputs() {
  // mediaTypes can be moved to local state of page.tsx, because were updating the form state
  // only if media type box exists

  const { setValue } = useFormContext<CollectorFormData>();

  return (
    <form
      id="MediaInputForm"
      className="MediaInputs flex flex-col items-center gap-4 sm:grid sm:grid-cols-2"
    >
      <TextInput
        variant="multiline"
        label={`Book Titles`}
        rows={5}
        onChange={(e) => {
          const titleSearchList = titleCollectionListConversion(e.target.value);
          setValue(`collectionList.book`, titleSearchList);
        }}
      />
    </form>
  );
}
