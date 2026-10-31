import titleCollectionListConversion from 'lib/helpers/title-collection-list-conversion';
import { MEDIA_TYPE_DEFINITIONS, MediaType } from 'lib/constants/media-types';
import type { MediaVisibilityMap } from '@/mediacollector/media-checkboxes';
import { useFormContext, useWatch } from 'react-hook-form';
import { CollectorFormData } from '@/mediacollector/collector-form/collector-form-schema';

interface MediaInputProps {
  visibility: MediaVisibilityMap;
}

export default function MediaInputs({ visibility }: MediaInputProps) {
  const { setValue } = useFormContext<CollectorFormData>();
  const collectionList = useWatch<CollectorFormData, 'collectionList'>({
    name: 'collectionList',
  });
  const visibleMediaTypes = MEDIA_TYPE_DEFINITIONS.filter(
    ({ mediaType }) => visibility[mediaType],
  );

  const handleMediaInputChange = (mediaType: MediaType, value: string) => {
    const titleSearchList = titleCollectionListConversion(value);
    setValue(`collectionList.${mediaType}`, titleSearchList);
  };

  return (
    <div className="flex flex-wrap gap-3">
      {visibleMediaTypes.map(({ mediaType, plural }) => {
        const titleCount = collectionList?.[mediaType]?.length ?? 0;
        return (
          <label
            key={mediaType}
            className="flex flex-[1_1_180px] flex-col gap-1.5 text-[13px] font-bold text-[#6b645d]"
          >
            {plural}
            <textarea
              rows={5}
              className="resize-y rounded-[10px] border border-[#ddd5cb] bg-[#fcfbf9] px-3 py-2.5 text-sm leading-normal text-[#2b2826]"
              onChange={(e) =>
                handleMediaInputChange(mediaType, e.target.value)
              }
            />
            <span className="text-xs font-semibold">
              {titleCount} {titleCount === 1 ? 'title' : 'titles'}
            </span>
          </label>
        );
      })}
    </div>
  );
}
