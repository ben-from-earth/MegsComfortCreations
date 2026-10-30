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
            <fieldset className="m-0 rounded-[14px] border border-[#e6e0d8] bg-[#faf8f5] p-2.5">
              <legend className="sr-only">Bleed size</legend>
              <div className="grid grid-cols-2 gap-2">
                {PNG_FORMAT_OPTIONS.map((option) => {
                  const isChecked = field.value === option.value;
                  return (
                    <label
                      key={option.value}
                      className={`relative flex min-h-16 cursor-pointer flex-col items-center justify-center rounded-[10px] border-[1.5px] bg-white has-focus-visible:outline-2 has-focus-visible:outline-[#3d3935] ${
                        isChecked
                          ? 'border-[#3d3935] bg-[#efebe5]'
                          : 'border-[#ddd5cb]'
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
                      <span className="text-xl font-extrabold">
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
