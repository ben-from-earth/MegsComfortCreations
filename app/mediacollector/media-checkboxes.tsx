import { MEDIA_TYPE_DEFINITIONS, MediaType } from 'lib/constants/media-types';

export type MediaVisibilityMap = Record<MediaType, boolean>;

interface MediaCheckboxesProps {
  visibility: MediaVisibilityMap;
  onToggle: (mediaType: MediaType) => void;
}

export default function MediaCheckboxes({
  visibility,
  onToggle,
}: MediaCheckboxesProps) {
  return (
    <div className="m-2 flex flex-row flex-wrap content-center justify-center gap-5">
      {MEDIA_TYPE_DEFINITIONS.map(({ mediaType, plural }) => (
        <label key={mediaType} className="text-3xl tracking-wider">
          <input
            checked={visibility[mediaType]}
            className="m-1.5"
            type="checkbox"
            onChange={() => onToggle(mediaType)}
          />
          {plural}
        </label>
      ))}
    </div>
  );
}
