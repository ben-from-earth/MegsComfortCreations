// react, redux imports
import { memo } from 'react';

// components
import CollectedCoversBlock from '@/mediacollector/CollectedCoversBlock';

import { useFormContext } from 'react-hook-form';
import { CollectorFormData } from './collector-form/collectorFormSchema';

const TitleBlockContainer = memo(function TitleBlockContainer({
  handleDeleteBlock,
  blockIdsWithErrors,
}: {
  handleDeleteBlock: (blockID: string) => void;
  blockIdsWithErrors: string[];
}) {
  const { watch } = useFormContext<CollectorFormData>();
  const blocks = watch('collectedData') ?? [];

  return (
    <div className="grid w-full grid-cols-1 gap-3 p-2 sm:grid-flow-dense sm:grid-cols-[repeat(auto-fill,minmax(30rem,1fr))] sm:auto-rows-[15rem]">
      {blocks.map((block, idx) => (
        <div
          key={block.blockID}
          className={block.type === 'book' ? 'sm:row-span-2' : 'sm:row-span-1'}
        >
          <CollectedCoversBlock
            index={idx}
            info={block}
            handleDeleteBlock={handleDeleteBlock}
            hasError={blockIdsWithErrors.includes(block.blockID)}
          />
        </div>
      ))}
    </div>
  );
});

export default TitleBlockContainer;
