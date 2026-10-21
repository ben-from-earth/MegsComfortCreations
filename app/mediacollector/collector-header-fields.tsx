import {
  FormControl,
  FormField,
  FormItem,
  FormMessage,
} from '@/components/ui/form';
import TextInput from '@/components/shared/text-input';
import {
  BOOK_CLUB_REPEAT_MAX,
  type CollectorFormData,
} from './collector-form/collector-form-schema';
import { useFormContext } from 'react-hook-form';

export default function CollectorHeaderFields() {
  const { control } = useFormContext<CollectorFormData>();

  return (
    <>
      <FormField
        control={control}
        name="customerName"
        render={({ field }) => (
          <FormItem>
            <FormControl>
              <TextInput
                label="Customer Full Name"
                variant="normal"
                name={field.name}
                value={field.value}
                onChange={field.onChange}
                onBlur={field.onBlur}
              />
            </FormControl>
          </FormItem>
        )}
      />
      <FormField
        control={control}
        name="orderNumber"
        render={({ field }) => (
          <FormItem>
            <FormControl>
              <TextInput
                label="Order Number"
                variant="normal"
                name={field.name}
                value={field.value}
                onChange={field.onChange}
                onBlur={field.onBlur}
              />
            </FormControl>
          </FormItem>
        )}
      />
      <FormField
        control={control}
        name="bookClubRepeat"
        render={({ field }) => (
          <FormItem className="flex flex-col items-center">
            <FormControl>
              <TextInput
                label="Book Club Repeat Number"
                variant="normal"
                name={field.name}
                value={String(field.value)}
                onChange={(event) => {
                  const parsed = Number(event.target.value);
                  if (!Number.isFinite(parsed)) {
                    field.onChange(0);
                    return;
                  }
                  field.onChange(Math.min(parsed, BOOK_CLUB_REPEAT_MAX));
                }}
                onBlur={field.onBlur}
              />
            </FormControl>
            <FormMessage />
          </FormItem>
        )}
      />
    </>
  );
}
