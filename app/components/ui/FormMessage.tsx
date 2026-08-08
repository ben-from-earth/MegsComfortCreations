'use client';

import {
  get,
  useFormContext,
  type FieldPath,
  type FieldValues,
} from 'react-hook-form';

export default function FormMessage<TFieldValues extends FieldValues>({
  name,
}: {
  name: FieldPath<TFieldValues>;
}) {
  const {
    formState: { errors },
  } = useFormContext<TFieldValues>();
  const error = get(errors, name);
  const message =
    typeof error?.message === 'string' ? error.message : undefined;

  if (!message) {
    return null;
  }

  return (
    <p className='m-0 font-["Just_Another_Hand"] text-2xl tracking-wider text-red-600'>
      {message}
    </p>
  );
}
