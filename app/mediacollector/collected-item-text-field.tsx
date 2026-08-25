import {
  FormControl,
  FormField,
  FormItem,
  FormLabel,
  FormMessage,
} from '@/components/ui/form';
import type { CollectorFormData } from '@/mediacollector/collector-form/collector-form-schema';
import { MediaType } from 'lib/constants/media-types';
import { useFormContext } from 'react-hook-form';

const NUMBER_FIELD_NAMES = ['pubYear', 'pageCount'] as const;

type CollectedTextFieldName = 'title' | 'author' | 'pubYear' | 'pageCount';

export interface CollectedItemTextFieldProps {
  name: CollectedTextFieldName;
  label: string;
  type: MediaType;
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
  type,
  index,
}: CollectedItemTextFieldProps) {
  const { control } = useFormContext<CollectorFormData>();
  const isNumberField = isNumberFieldName(name);
  const inputStyling =
    type !== 'book'
      ? { marginBottom: '20px', width: '200px' }
      : { width: '300px' };

  return (
    <FormField
      control={control}
      name={`collectedData.${index}.blockInfo.${name}`}
      render={({ field }) => (
        <FormItem>
          <div className="relative">
            <FormLabel className="absolute right-full mr-2 w-fit translate-y-1/8 text-right text-3xl text-nowrap">
              {label}:
            </FormLabel>
            <FormControl>
              <textarea
                className="content-center rounded-sm bg-white pl-2 font-[Arial] text-black"
                style={inputStyling}
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
          </div>
          <FormMessage />
        </FormItem>
      )}
    />
  );
}
