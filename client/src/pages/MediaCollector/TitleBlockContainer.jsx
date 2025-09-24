import CollectedCoversBlock from '@/pages/MediaCollector/CollectedCoversBlock';
import { memo } from 'react';

const TitleBlockContainer = memo(function ({ blocks, handleDeleteBlock }) {
  return (
    <div className="flex w-full flex-row flex-wrap">
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
