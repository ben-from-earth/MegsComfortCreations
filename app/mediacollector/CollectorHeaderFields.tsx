import {
  FormControl,
  FormField,
  FormItem,
  FormMessage,
} from '@/components/ui/form';
import TextInput from '@/components/shared/TextInput';
import type { CollectorFormData } from './collector-form/collectorFormSchema';

export default function CollectorHeaderFields() {
  return (
    <>
      <FormField<CollectorFormData, 'customerName'>
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
      <FormField<CollectorFormData, 'orderNumber'>
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
      <FormField<CollectorFormData, 'bookClubRepeat'>
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
                  field.onChange(Number(event.target.value));
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
