import { useWatch } from 'react-hook-form';
import { MEDIA_TYPE_DEFINITIONS, MediaType } from 'lib/constants/media-types';
import type { CollectorFormData } from './collector-form/collector-form-schema';

export type MediaVisibilityMap = Record<MediaType, boolean>;

interface MediaCheckboxesProps {
  visibility: MediaVisibilityMap;
  onToggle: (mediaType: MediaType) => void;
}

export default function MediaCheckboxes({
  visibility,
  onToggle,
}: MediaCheckboxesProps) {
  const collectionList = useWatch<CollectorFormData, 'collectionList'>({
    name: 'collectionList',
  });

  return (
    <div className="flex flex-wrap gap-2">
      {MEDIA_TYPE_DEFINITIONS.map(({ mediaType, plural }) => {
        const isVisible = visibility[mediaType];
        const titleCount = collectionList?.[mediaType]?.length ?? 0;
        return (
          <button
            key={mediaType}
            type="button"
            aria-pressed={isVisible}
            onClick={() => onToggle(mediaType)}
            className={`h-9 cursor-pointer rounded-full border-2 px-3.5 text-[13px] font-extrabold ${
              isVisible
                ? 'border-[#2a4637] bg-[#efebe5] text-[#2b2826]'
                : 'border-[#e0d9d0] bg-white text-[#6b645d]'
            }`}
          >
            {isVisible && titleCount > 0 ? `${plural} · ${titleCount}` : plural}
          </button>
        );
      })}
    </div>
  );
}
