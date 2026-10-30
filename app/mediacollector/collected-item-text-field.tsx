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

export interface CollectedItemTextFieldProps {
  name: CollectedTextFieldName;
  label: string;
  index: number;
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
              className={`h-9.5 min-w-0 rounded-lg border border-[#e0d9d0] bg-[#fcfbf9] px-2.5 text-sm text-[#2b2826] aria-invalid:border-red-600 ${
                name === 'title' ? 'font-bold' : 'font-normal'
              }`}
              name={field.name}
              value={field.value == null ? '' : String(field.value)}
              onBlur={field.onBlur}
              onChange={(event) => {
                if (isNumberField) {
                  const parsed = Number(event.target.value);
                  field.onChange(
                    event.target.value.trim() === '' ||
                      !Number.isFinite(parsed)
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
