import { useFormContext } from 'react-hook-form';
import {
  FormControl,
  FormField,
  FormItem,
  FormLabel,
  FormMessage,
} from '@/components/ui/form';
import type { MediaItemForm } from '@/mediacollector/collector-form/media-item-form-schema';
import {
  editStateClassNames,
  type CollectedFieldEditState,
} from '@/mediacollector/collected-item-text-field';

type LibraryTextFieldName = 'title' | 'author' | 'pubYear' | 'pageCount';

export default function LibraryItemTextField({
  name,
  label,
  editState,
}: {
  name: LibraryTextFieldName;
  label: string;
  editState?: CollectedFieldEditState;
}) {
  const { control } = useFormContext<MediaItemForm>();
  const isNumberField = name === 'pubYear' || name === 'pageCount';

  return (
    <FormField
      control={control}
      name={`blockInfo.${name}`}
      render={({ field }) => (
        <FormItem className="flex min-w-0 flex-col gap-1 text-xs font-bold text-[#6b645d]">
          <FormLabel>{label}</FormLabel>
          <FormControl>
            <input
              type="text"
              inputMode={isNumberField ? 'numeric' : undefined}
              readOnly={editState === 'readOnly'}
              className={`h-9.5 min-w-0 rounded-lg border px-2.5 text-sm text-[#2b2826] aria-invalid:border-red-600 ${
                editStateClassNames[editState ?? 'editable']
              } ${name === 'title' ? 'font-bold' : 'font-normal'}`}
              name={field.name}
              value={field.value == null ? '' : String(field.value)}
              onBlur={field.onBlur}
              onChange={(event) => {
                if (isNumberField) {
                  const parsed = Number(event.target.value);
                  field.onChange(
                    event.target.value.trim() === '' || !Number.isFinite(parsed)
                      ? null
                      : parsed,
                  );
                  return;
                }
                field.onChange(event.target.value);
              }}
            />
          </FormControl>
          <FormMessage className="m-0 text-xs font-bold text-red-600" />
        </FormItem>
      )}
    />
  );
}
