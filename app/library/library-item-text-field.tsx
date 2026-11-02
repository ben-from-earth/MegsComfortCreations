import { useFormContext } from 'react-hook-form';
import { FormField } from '@/components/ui/form';
import type { MediaItemForm } from '@/mediacollector/collector-form/media-item-form-schema';
import {
  MediaTextFieldControl,
  type CollectedFieldEditState,
  type MediaTextFieldName,
} from '@/mediacollector/collected-item-text-field';

export default function LibraryItemTextField({
  name,
  label,
  editState,
}: {
  name: MediaTextFieldName;
  label: string;
  editState?: CollectedFieldEditState;
}) {
  const { control } = useFormContext<MediaItemForm>();

  return (
    <FormField
      control={control}
      name={`blockInfo.${name}`}
      render={({ field }) => (
        <MediaTextFieldControl
          name={name}
          label={label}
          editState={editState}
          field={field}
        />
      )}
    />
  );
}
