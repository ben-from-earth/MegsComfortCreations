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
        <FormItem className="flex flex-col items-center">
          <FormControl>
            <div className="flex gap-2.5">
              {PNG_FORMAT_OPTIONS.map((option) => (
                <label
                  key={option.value}
                  className='font-["Just_Another_Hand"] text-2xl tracking-wider'
                >
                  <input
                    className="m-2"
                    type="radio"
                    name={field.name}
                    value={option.value}
                    checked={field.value === option.value}
                    onChange={() => field.onChange(option.value)}
                    onBlur={field.onBlur}
                  />
                  {option.label}
                </label>
              ))}
            </div>
          </FormControl>
          <FormMessage />
        </FormItem>
      )}
    />
  );
}
