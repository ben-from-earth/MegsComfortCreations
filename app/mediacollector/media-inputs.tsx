import titleCollectionListConversion from 'lib/helpers/title-collection-list-conversion';
import { MEDIA_TYPE_DEFINITIONS, MediaType } from 'lib/constants/media-types';
import type { MediaVisibilityMap } from '@/mediacollector/media-checkboxes';
import { useId } from 'react';
import { useFormContext, useWatch } from 'react-hook-form';
import { CollectorFormData } from '@/mediacollector/collector-form/collector-form-schema';
import MovieFormatToggle from '@/mediacollector/movie-format-toggle';

interface MediaInputProps {
  visibility: MediaVisibilityMap;
}

export default function MediaInputs({ visibility }: MediaInputProps) {
  const inputIdPrefix = useId();
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
        const textareaId = `${inputIdPrefix}-${mediaType}`;
        return (
          <div
            key={mediaType}
            className="flex flex-[1_1_180px] flex-col gap-1.5 text-[13px] font-bold text-[#6b645d]"
          >
            <div className="flex min-h-6 items-center justify-between gap-2">
              <label htmlFor={textareaId}>{plural}</label>
              {mediaType === 'movie' ? <MovieFormatToggle /> : null}
            </div>
            <textarea
              id={textareaId}
              rows={5}
              className="resize-y rounded-[10px] border border-[#ddd5cb] bg-[#fcfbf9] px-3 py-2.5 text-sm leading-normal text-[#2b2826]"
              onChange={(e) =>
                handleMediaInputChange(mediaType, e.target.value)
              }
            />
            <span className="text-xs font-semibold">
              {titleCount} {titleCount === 1 ? 'title' : 'titles'}
            </span>
          </div>
        );
      })}
    </div>
  );
}
