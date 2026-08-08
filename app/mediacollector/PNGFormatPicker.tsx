import { useFormContext } from 'react-hook-form';
import { CollectorFormData } from './collector-form/collectorFormSchema';
import FormMessage from '@/components/ui/FormMessage';

export default function PNGFormatPicker() {
  const { watch, setValue } = useFormContext<CollectorFormData>();

  return (
    <div className="flex flex-col items-center">
      <div className="flex gap-2.5">
        <label className='font-["Just_Another_Hand"] text-2xl tracking-wider'>
          <input
            className="m-2"
            id={'3mm'}
            type="checkbox"
            checked={watch('pngFormat') === '3'}
            onChange={(e) => {
              if (e.target.checked === true) {
                setValue('pngFormat', '3', { shouldValidate: true });
              } else {
                setValue('pngFormat', null, { shouldValidate: true });
              }
            }}
          />
          3mm PNG Format
        </label>
        <label className='font-["Just_Another_Hand"] text-2xl tracking-wider'>
          <input
            className="m-2"
            id={'5mm'}
            type="checkbox"
            checked={watch('pngFormat') === '5'}
            onChange={(e) => {
              if (e.target.checked === true) {
                setValue('pngFormat', '5', { shouldValidate: true });
              } else {
                setValue('pngFormat', null, { shouldValidate: true });
              }
            }}
          />
          5mm PNG Format
        </label>
      </div>
      <FormMessage<CollectorFormData> name="pngFormat" />
    </div>
  );
}
