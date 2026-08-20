import {
  FormControl,
  FormField,
  FormItem,
  useFormField,
} from '@/components/ui/form';
import type { CollectorFormData } from '@/mediacollector/collector-form/collectorFormSchema';
import { MediaType } from 'lib/constants/mediaTypes';
import type { ControllerRenderProps } from 'react-hook-form';

const NUMBER_FIELD_NAMES = ['pubYear', 'pageCount'] as const;

type CollectedTextFieldName = 'title' | 'author' | 'pubYear' | 'pageCount';

export interface MyTextAreaProps {
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

function CollectedItemTextAreaControl({
  label,
  type,
  name,
  field,
}: {
  label: string;
  type: MediaType;
  name: CollectedTextFieldName;
  field: ControllerRenderProps<
    CollectorFormData,
    `collectedData.${number}.blockInfo.${CollectedTextFieldName}`
  >;
}) {
  const { formItemId } = useFormField();
  const isNumberField = isNumberFieldName(name);

  const inputStyling =
    type !== 'book'
      ? { marginBottom: '20px', width: '200px' }
      : { width: '300px' };
  const onChange = (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    if (isNumberField) {
      const parsed = Number(event.target.value);
      field.onChange(
        event.target.value.trim() === '' || !Number.isFinite(parsed)
          ? null
          : parsed,
      );
    }
    field.onChange(event.target.value);
  };

  return (
    <div className="relative">
      <label
        className="absolute right-full mr-2 w-fit translate-y-1/8 text-right text-3xl text-nowrap"
        htmlFor={formItemId}
      >
        {label}:
      </label>
      <FormControl>
        <textarea
          className="content-center rounded-sm bg-white pl-2 font-[Arial] text-black"
          style={inputStyling}
          name={field.name}
          value={field.value == null ? '' : String(field.value)}
          onBlur={field.onBlur}
          onChange={onChange}
        />
      </FormControl>
    </div>
  );
}

export default function MyTextArea({
  name,
  label,
  type,
  index,
}: MyTextAreaProps) {
  return (
    <FormField<
      CollectorFormData,
      `collectedData.${number}.blockInfo.${CollectedTextFieldName}`
    >
      name={`collectedData.${index}.blockInfo.${name}`}
      render={({ field }) => (
        <FormItem>
          <CollectedItemTextAreaControl
            label={label}
            type={type}
            name={name}
            field={field}
          />
        </FormItem>
      )}
    />
  );
}
