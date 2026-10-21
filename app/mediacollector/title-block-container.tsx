import { memo } from 'react';
import type { FieldArrayWithId } from 'react-hook-form';
import CollectedCoversBlock from '@/mediacollector/collected-covers-block';
import { CollectorFormData } from './collector-form/collector-form-schema';

const TitleBlockContainer = memo(function TitleBlockContainer({
  fields,
  onDelete,
  blockIdsWithErrors,
}: {
  fields: FieldArrayWithId<CollectorFormData, 'collectedData', 'fieldId'>[];
  onDelete: (index: number) => void;
  blockIdsWithErrors: string[];
}) {
  return (
    <div className="grid w-full grid-cols-1 gap-3 p-2 sm:grid-flow-dense sm:grid-cols-[repeat(auto-fill,minmax(30rem,1fr))] sm:auto-rows-[15rem]">
      {fields.map((field, index) => (
        <div
          key={field.fieldId}
          className={field.type === 'book' ? 'sm:row-span-2' : 'sm:row-span-1'}
        >
          <CollectedCoversBlock
            index={index}
            type={field.type}
            isDatabase={field.isDatabase}
            blockID={field.blockID}
            spineColor={field.blockInfo.spineColor}
            onDelete={onDelete}
            hasError={blockIdsWithErrors.includes(field.blockID)}
          />
        </div>
      ))}
    </div>
  );
});

export default TitleBlockContainer;
