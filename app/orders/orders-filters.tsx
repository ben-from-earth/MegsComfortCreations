import {
  MEDIA_TYPE_DEFINITIONS,
  type MediaType,
} from 'lib/constants/media-types';

export type OrdersFilterText = {
  name: string;
  number: string;
  title: string;
  min: string;
  max: string;
};

const fieldClassName =
  'h-10 w-full rounded-[10px] border border-[#ddd5cb] bg-[#fcfbf9] px-3 text-sm text-[#2b2826] focus-visible:outline-3 focus-visible:outline-offset-2 focus-visible:outline-[#3d3935]';
const labelClassName =
  'flex flex-col gap-1 text-[11px] font-extrabold tracking-[0.6px] text-[#6b645d] uppercase';

export default function OrdersFilters({
  text,
  onTextChange,
  type,
  onTypeChange,
}: {
  text: OrdersFilterText;
  onTextChange: (key: keyof OrdersFilterText, value: string) => void;
  type: MediaType | null;
  onTypeChange: (type: MediaType | null) => void;
}) {
  return (
    <div className="flex flex-wrap items-end gap-3 rounded-2xl border border-[#e6e0d8] bg-white px-[18px] py-4">
      <label className={`${labelClassName} min-w-[200px] flex-1`}>
        Customer name
        <input
          type="search"
          value={text.name}
          onChange={(event) => onTextChange('name', event.target.value)}
          className={fieldClassName}
        />
      </label>
      <label className={`${labelClassName} w-[160px]`}>
        Order number
        <input
          type="search"
          value={text.number}
          onChange={(event) => onTextChange('number', event.target.value)}
          className={fieldClassName}
        />
      </label>
      <label className={`${labelClassName} min-w-[200px] flex-1`}>
        Contains title
        <input
          type="search"
          value={text.title}
          onChange={(event) => onTextChange('title', event.target.value)}
          className={fieldClassName}
        />
      </label>
      <label className={`${labelClassName} w-[150px]`}>
        Media type
        <select
          value={type ?? ''}
          onChange={(event) =>
            onTypeChange(
              MEDIA_TYPE_DEFINITIONS.find(
                (definition) => definition.mediaType === event.target.value,
              )?.mediaType ?? null,
            )
          }
          className={fieldClassName}
        >
          <option value="">Any</option>
          {MEDIA_TYPE_DEFINITIONS.map((definition) => (
            <option key={definition.mediaType} value={definition.mediaType}>
              {definition.plural}
            </option>
          ))}
        </select>
      </label>
      <label className={`${labelClassName} w-[90px]`}>
        Min count
        <input
          type="number"
          min={0}
          value={text.min}
          disabled={type === null}
          onChange={(event) => onTextChange('min', event.target.value)}
          className={`${fieldClassName} disabled:opacity-50`}
        />
      </label>
      <label className={`${labelClassName} w-[90px]`}>
        Max count
        <input
          type="number"
          min={0}
          value={text.max}
          disabled={type === null}
          onChange={(event) => onTextChange('max', event.target.value)}
          className={`${fieldClassName} disabled:opacity-50`}
        />
      </label>
    </div>
  );
}
