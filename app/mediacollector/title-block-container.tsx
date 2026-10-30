import { memo } from 'react';
import type { FieldArrayWithId } from 'react-hook-form';
import CollectedCoversBlock from '@/mediacollector/collected-covers-block';
import { CollectorFormData } from './collector-form/collector-form-schema';

export type ResultsFilter = 'all' | 'new' | 'database';

const TitleBlockContainer = memo(function TitleBlockContainer({
  fields,
  filter,
  onDelete,
  blockIdsWithErrors,
}: {
  fields: FieldArrayWithId<CollectorFormData, 'collectedData', 'fieldId'>[];
  filter: ResultsFilter;
  onDelete: (index: number) => void;
  blockIdsWithErrors: string[];
}) {
  return (
    <div className="grid grid-cols-[repeat(auto-fill,minmax(min(380px,100%),1fr))] gap-5">
      {fields.map((field, index) => {
        if (
          (filter === 'new' && field.isDatabase) ||
          (filter === 'database' && !field.isDatabase)
        ) {
          return null;
        }
        return (
          <CollectedCoversBlock
            key={field.fieldId}
            index={index}
            type={field.type}
            isDatabase={field.isDatabase}
            blockID={field.blockID}
            onDelete={onDelete}
            hasError={blockIdsWithErrors.includes(field.blockID)}
          />
        );
      })}
    </div>
  );
});

export default TitleBlockContainer;
