import {
  FormControl,
  FormField,
  FormItem,
  FormMessage,
} from '@/components/ui/form';
import {
  PNG_FORMAT_OPTIONS,
  type CollectorFormData,
} from './collector-form/collector-form-schema';
import { useFormContext } from 'react-hook-form';

export default function PNGFormatPicker() {
  const { control } = useFormContext<CollectorFormData>();

  return (
    <FormField
      control={control}
      name="pngFormat"
      render={({ field }) => (
        <FormItem className="flex flex-col gap-1.5">
          <FormControl>
            <fieldset className="m-0 border-0 p-0">
              <legend className="sr-only">Bleed size</legend>
              <div className="grid grid-cols-2 gap-2">
                {PNG_FORMAT_OPTIONS.map((option) => {
                  const isChecked = field.value === option.value;
                  return (
                    <label
                      key={option.value}
                      className={`relative flex h-11 cursor-pointer items-center justify-center rounded-[10px] border-[1.5px] has-focus-visible:outline-2 has-focus-visible:outline-offset-2 has-focus-visible:outline-[#2a4637] ${
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
                        onClick={() => {
                          if (isChecked) {
                            field.onChange(null);
                          }
                        }}
                        onBlur={field.onBlur}
                      />
                      <span className="text-lg font-extrabold">
                        {option.label}
                      </span>
                    </label>
                  );
                })}
              </div>
            </fieldset>
          </FormControl>
          <FormMessage className="m-0 text-xs font-bold text-red-600" />
        </FormItem>
      )}
    />
  );
}
