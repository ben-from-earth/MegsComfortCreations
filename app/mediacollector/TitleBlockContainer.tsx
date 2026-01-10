// react, redux imports
import { memo } from 'react';

// components
import CollectedCoversBlock from '@/mediacollector/CollectedCoversBlock';

import { useFormContext } from 'react-hook-form';
import { CollectorFormData } from './collector-form/collectorFormSchema';

const TitleBlockContainer = memo(function TitleBlockContainer({
  handleDeleteBlock,
}: {
  handleDeleteBlock: (blockID: string) => void;
}) {
  const { watch } = useFormContext<CollectorFormData>();
  const blocks = watch('collectedData');
  return (
    <div className="flex w-full flex-row flex-wrap gap-3 p-2">
      {blocks.map((block, idx) => (
        <CollectedCoversBlock
          index={idx}
          info={block}
          key={block.blockID}
          handleDeleteBlock={handleDeleteBlock}
        />
      ))}
    </div>
  );
});

export default TitleBlockContainer;
