// interfaces and types
import { MediaType } from 'lib/constants/mediaTypes';
import { CollectorFormData } from '@/mediacollector/collector-form/collectorFormSchema';

import { useFormContext } from 'react-hook-form';

export interface MyTextAreaProps {
  name: 'title' | 'author' | 'pubYear' | 'pageCount';
  label: string;
  type: MediaType;
  blockID: string;
  value: string | number;
}

export default function MyTextArea({
  name,
  label,
  type,
  blockID,
  value,
}: MyTextAreaProps) {
  const { watch, setValue } = useFormContext<CollectorFormData>();
  const collectedData = watch('collectedData');
  const block = collectedData.find((block) => block.blockID === blockID);
  if (!block) {
    return null;
  }

  return (
    <div className="relative">
      <label
        className="absolute right-full mr-2 w-fit translate-y-1/8 text-right text-3xl text-nowrap"
        htmlFor={name}
      >
        {label}:
      </label>
      <textarea
        className="content-center rounded-sm bg-white pl-2 font-[Arial] text-black"
        style={
          type !== 'book'
            ? { marginBottom: '20px', width: '200px' }
            : { width: '300px' }
        }
        name={name}
        defaultValue={value}
        onChange={(e) => {
          const newText = e.target.value;
          const newBlock = {
            ...block,
            blockInfo: { ...block?.blockInfo, [name]: newText },
          };
          setValue(
            'collectedData',
            collectedData.map((b) => (b.blockID === blockID ? newBlock : b)),
          );
        }}
      ></textarea>
    </div>
  );
}
