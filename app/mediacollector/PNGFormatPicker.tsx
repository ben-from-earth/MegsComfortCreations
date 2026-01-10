// react, redux imports
import { Dispatch, SetStateAction } from 'react';
import { useFormContext } from 'react-hook-form';
import { CollectorFormData } from './collector-form/collectorFormSchema';

// interfaces and types

interface PNGFormatPickerProps {
  pngError: boolean;
  setPNGError: Dispatch<SetStateAction<boolean>>;
}

export default function PNGFormatPicker({
  pngError,
  setPNGError,
}: PNGFormatPickerProps) {
  const { watch, setValue } = useFormContext<CollectorFormData>();
  return (
    <div className="flex flex-col items-center">
      <p
        className='m-0 font-["Just_Another_Hand"] text-2xl tracking-wider'
        style={{
          visibility: pngError ? 'visible' : 'hidden',
          color: 'red',
        }}
      >
        Please select a PNG template option
      </p>
      <div className="flex gap-2.5">
        <label className='font-["Just_Another_Hand"] text-2xl tracking-wider'>
          <input
            className="m-2"
            id={'3mm'}
            type="checkbox"
            checked={watch('pngFormat') === '3'}
            onChange={(e) => {
              if (e.target.checked === true) {
                setPNGError(false);
                setValue('pngFormat', '3');
              } else {
                setValue('pngFormat', undefined);
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
                setValue('pngFormat', '5');
                setPNGError(false);
              } else {
                setValue('pngFormat', undefined);
              }
            }}
          />
          5mm PNG Format
        </label>
      </div>
    </div>
  );
}
