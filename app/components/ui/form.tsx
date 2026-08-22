'use client';

import { cloneElement, createContext, use, useId } from 'react';
import type { ComponentProps, ReactElement } from 'react';
import type { ControllerProps, FieldPath, FieldValues } from 'react-hook-form';
import {
  Controller,
  FormProvider,
  useFormContext,
  useFormState,
} from 'react-hook-form';

const Form = FormProvider;

interface FormFieldContextValue<
  TFieldValues extends FieldValues = FieldValues,
  TName extends FieldPath<TFieldValues> = FieldPath<TFieldValues>,
> {
  name: TName;
}

const FormFieldContext = createContext<FormFieldContextValue | null>(null);

function FormField<
  TFieldValues extends FieldValues = FieldValues,
  TName extends FieldPath<TFieldValues> = FieldPath<TFieldValues>,
>(props: ControllerProps<TFieldValues, TName>) {
  return (
    <FormFieldContext value={{ name: props.name }}>
      <Controller {...props} />
    </FormFieldContext>
  );
}

interface FormItemContextValue {
  id: string;
}

const FormItemContext = createContext<FormItemContextValue | null>(null);

function useFormField() {
  const fieldContext = use(FormFieldContext);
  const itemContext = use(FormItemContext);

  if (!fieldContext) {
    throw new Error('useFormField should be used within <FormField>');
  }
  if (!itemContext) {
    throw new Error('useFormField should be used within <FormItem>');
  }

  const { getFieldState } = useFormContext();
  const formState = useFormState({ name: fieldContext.name });
  const fieldState = getFieldState(fieldContext.name, formState);

  return {
    id: itemContext.id,
    name: fieldContext.name,
    formItemId: `${itemContext.id}-form-item`,
    formMessageId: `${itemContext.id}-form-item-message`,
    ...fieldState,
  };
}

function FormItem({
  className,
  ...props
}: ComponentProps<'div'>) {
  const id = useId();

  return (
    <FormItemContext value={{ id }}>
      <div className={className} {...props} />
    </FormItemContext>
  );
}

function FormLabel({
  className,
  ...props
}: ComponentProps<'label'>) {
  const { formItemId } = useFormField();

  return <label className={className} htmlFor={formItemId} {...props} />;
}

function FormControl({
  children,
}: {
  children: ReactElement<Record<string, unknown>>;
}) {
  const { error, formItemId, formMessageId } = useFormField();

  return cloneElement(children, {
    id: formItemId,
    'aria-invalid': !!error,
    'aria-describedby': error ? formMessageId : undefined,
  });
}

function FormMessage({ className }: { className?: string }) {
  const { error, formMessageId } = useFormField();
  const message = typeof error?.message === 'string' ? error.message : undefined;

  if (!message) {
    return null;
  }

  return (
    <p
      id={formMessageId}
      className={
        className ??
        'm-0 font-["Just_Another_Hand"] text-2xl tracking-wider text-red-600'
      }
    >
      {message}
    </p>
  );
}

export {
  Form,
  FormControl,
  FormField,
  FormItem,
  FormLabel,
  FormMessage,
  useFormField,
};
