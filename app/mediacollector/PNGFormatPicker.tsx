// react, redux imports
import { Dispatch, SetStateAction } from 'react';

// interfaces and types

interface PNGFormatPickerProps {
  pngTemplateChecks: boolean[];
  pngError: boolean;
  setPNGError: Dispatch<SetStateAction<boolean>>;
  setPNGTemplate: Dispatch<SetStateAction<number | undefined>>;
  setPNGTemplateChecks: Dispatch<SetStateAction<boolean[]>>;
}

export default function PNGFormatPicker({
  pngTemplateChecks,
  pngError,
  setPNGError,
  setPNGTemplate,
  setPNGTemplateChecks,
}: PNGFormatPickerProps) {
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
            checked={pngTemplateChecks[0]}
            onChange={(e) => {
              if (e.target.checked === true) {
                setPNGError(false);
                setPNGTemplateChecks([true, false]);
                setPNGTemplate(3);
              } else {
                setPNGTemplateChecks((prev) => [false, prev[1]]);
                setPNGTemplate(undefined);
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
            checked={pngTemplateChecks[1]}
            onChange={(e) => {
              if (e.target.checked === true) {
                setPNGTemplateChecks([false, true]);
                setPNGTemplate(5);
                setPNGError(false);
              } else {
                setPNGTemplateChecks((prev) => [prev[0], false]);
                setPNGTemplate(undefined);
              }
            }}
          />
          5mm PNG Format
        </label>
      </div>
    </div>
  );
}
