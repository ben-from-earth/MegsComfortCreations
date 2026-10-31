import {
  FormControl,
  FormField,
  FormItem,
  FormLabel,
  FormMessage,
} from '@/components/ui/form';
import {
  BOOK_CLUB_REPEAT_MAX,
  BOOK_CLUB_REPEAT_MIN,
  type CollectorFormData,
} from './collector-form/collector-form-schema';
import { useFormContext } from 'react-hook-form';

const labelClassName = 'flex flex-col gap-1.5 text-[13px] font-bold text-[#6b645d]';
const inputClassName =
  'h-11 rounded-[10px] border border-[#ddd5cb] bg-[#fcfbf9] px-3 text-[15px] text-[#2b2826] aria-invalid:border-red-600';
const stepperButtonClassName =
  'h-11 w-11 cursor-pointer rounded-[10px] border border-[#ddd5cb] bg-white text-xl text-[#2b2826] disabled:cursor-not-allowed disabled:opacity-40';
const errorClassName = 'm-0 text-xs font-bold text-red-600';

export default function CollectorHeaderFields() {
  const { control } = useFormContext<CollectorFormData>();

  return (
    <>
      <FormField
        control={control}
        name="customerName"
        render={({ field }) => (
          <FormItem className={labelClassName}>
            <FormLabel>Customer full name</FormLabel>
            <FormControl>
              <input
                type="text"
                className={inputClassName}
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
          <FormItem className={labelClassName}>
            <FormLabel>Etsy order number</FormLabel>
            <FormControl>
              <input
                type="text"
                className={inputClassName}
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
          <FormItem className="flex flex-col items-center gap-1.5">
            <FormLabel className="text-[13px] font-bold text-[#6b645d]">
              Book club repeat
            </FormLabel>
            <div className="flex items-center justify-center gap-2">
              <button
                type="button"
                aria-label="Decrease repeat"
                className={stepperButtonClassName}
                disabled={field.value <= BOOK_CLUB_REPEAT_MIN}
                onClick={() =>
                  field.onChange(
                    Math.max(field.value - 1, BOOK_CLUB_REPEAT_MIN),
                  )
                }
              >
                −
              </button>
              <FormControl>
                <input
                  type="number"
                  min={BOOK_CLUB_REPEAT_MIN}
                  max={BOOK_CLUB_REPEAT_MAX}
                  className="h-11 w-18 [appearance:textfield] rounded-[10px] border border-[#ddd5cb] bg-[#fcfbf9] text-center text-lg font-extrabold text-[#2b2826] aria-invalid:border-red-600 [&::-webkit-inner-spin-button]:appearance-none [&::-webkit-outer-spin-button]:appearance-none"
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
              <button
                type="button"
                aria-label="Increase repeat"
                className={stepperButtonClassName}
                disabled={field.value >= BOOK_CLUB_REPEAT_MAX}
                onClick={() =>
                  field.onChange(
                    Math.min(
                      Math.max(field.value + 1, BOOK_CLUB_REPEAT_MIN),
                      BOOK_CLUB_REPEAT_MAX,
                    ),
                  )
                }
              >
                +
              </button>
            </div>
            <FormMessage className={errorClassName} />
          </FormItem>
        )}
      />
    </>
  );
}
