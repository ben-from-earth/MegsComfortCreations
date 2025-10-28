// react, redux imports
import { memo } from 'react';

// components
import CollectedCoversBlock from '@/app/mediacollector/CollectedCoversBlock';

// necessary imports from collector state slice
import { collectedBlockInformation } from '@/lib/state/slices/collectorSlice';

// interfaces and types
import { MediaType } from '@/lib/interfaces/globalInterfaces';

const TitleBlockContainer = memo(function TitleBlockContainer({
  blocks,
  handleDeleteBlock,
}: {
  blocks: collectedBlockInformation[];
  handleDeleteBlock: (
    blockID: string,
    type: MediaType,
    deleteBlock: boolean,
    urls: string[],
  ) => void;
}) {
  return (
    <div className="flex w-full flex-row flex-wrap gap-3 p-2">
      {blocks.map((b) => (
        <CollectedCoversBlock
          info={b}
          key={b.blockID}
          handleDeleteBlock={handleDeleteBlock}
        />
      ))}
    </div>
  );
});

export default TitleBlockContainer;
