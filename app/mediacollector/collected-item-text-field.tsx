import {
  FormControl,
  FormField,
  FormItem,
  FormLabel,
  FormMessage,
} from '@/components/ui/form';
import type { CollectorFormData } from '@/mediacollector/collector-form/collector-form-schema';
import { useFormContext } from 'react-hook-form';

const NUMBER_FIELD_NAMES = ['pubYear', 'pageCount'] as const;

type CollectedTextFieldName = 'title' | 'author' | 'pubYear' | 'pageCount';

export type CollectedFieldEditState = 'readOnly' | 'changed';

const editStateClassNames: Record<
  CollectedFieldEditState | 'editable',
  string
> = {
  editable: 'border-[#e0d9d0] bg-[#fcfbf9]',
  readOnly: 'border-[#ece6de] bg-[#f4f1ec]',
  changed: 'border-[#c9973f] bg-[#fffaf0] shadow-[inset_0_0_0_0.5px_#c9973f]',
};

export interface CollectedItemTextFieldProps {
  name: CollectedTextFieldName;
  label: string;
  index: number;
  editState?: CollectedFieldEditState;
}

function isNumberFieldName(
  name: CollectedTextFieldName,
): name is (typeof NUMBER_FIELD_NAMES)[number] {
  return NUMBER_FIELD_NAMES.some((fieldName) => fieldName === name);
}

export default function CollectedItemTextField({
  name,
  label,
  index,
  editState,
}: CollectedItemTextFieldProps) {
  const { control } = useFormContext<CollectorFormData>();
  const isNumberField = isNumberFieldName(name);

  return (
    <FormField
      control={control}
      name={`collectedData.${index}.blockInfo.${name}`}
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
