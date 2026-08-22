// interfaces and types
import titleCollectionListConversion from 'lib/helpers/titleCollectionListConversion';
import { MEDIA_TYPE_DEFINITIONS, MediaType } from 'lib/constants/mediaTypes';
import type { MediaVisibilityMap } from '@/mediacollector/MediaCheckboxes';

// component imports
import TextInput from '@/components/shared/TextInput';
import { useFormContext } from 'react-hook-form';
import { CollectorFormData } from '@/mediacollector/collector-form/collectorFormSchema';

interface MediaInputProps {
  visibility: MediaVisibilityMap;
}

export default function MediaInputs({ visibility }: MediaInputProps) {
  const { setValue } = useFormContext<CollectorFormData>();
  const visibleMediaTypes = MEDIA_TYPE_DEFINITIONS.filter(
    ({ mediaType }) => visibility[mediaType],
  );

  const handleMediaInputChange = (mediaType: MediaType, value: string) => {
    const titleSearchList = titleCollectionListConversion(value);
    setValue(`collectionList.${mediaType}`, titleSearchList);
  };

  return (
    <div
      id="MediaInputForm"
      className="MediaInputs flex flex-col items-center gap-4 sm:grid sm:grid-cols-2"
    >
      {visibleMediaTypes.map(({ mediaType, label }) => (
        <TextInput
          key={mediaType}
          variant="multiline"
          label={`${label} Titles`}
          rows={5}
          onChange={(e) => handleMediaInputChange(mediaType, e.target.value)}
        />
      ))}
    </div>
  );
}
