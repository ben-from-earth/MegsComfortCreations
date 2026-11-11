import { useController, useFormContext } from 'react-hook-form';
import { MOVIE_FORMAT_OPTIONS } from 'lib/constants/movie-formats';
import type { CollectorFormData } from './collector-form/collector-form-schema';

export default function MovieFormatToggle() {
  const { control } = useFormContext<CollectorFormData>();
  const { field } = useController({ control, name: 'movieFormat' });

  return (
    <fieldset className="m-0 border-0 p-0">
      <legend className="sr-only">Movie format</legend>
      <div className="flex gap-1">
        {MOVIE_FORMAT_OPTIONS.map((option) => {
          const isChecked = field.value === option.value;
          return (
            <label
              key={option.value}
              className={`flex h-6 cursor-pointer items-center rounded-md border px-2 text-xs font-bold has-focus-visible:outline-2 has-focus-visible:outline-offset-1 has-focus-visible:outline-[#2a4637] ${
                isChecked
                  ? 'border-[#2a4637] bg-[#2a4637] text-white'
                  : 'border-[#ddd5cb] bg-white text-[#2b2826]'
              }`}
            >
              <input
                className="sr-only"
                type="radio"
                name={field.name}
                value={option.value}
                checked={isChecked}
                onChange={() => field.onChange(option.value)}
                onBlur={field.onBlur}
              />
              {option.label}
            </label>
          );
        })}
      </div>
    </fieldset>
  );
}
