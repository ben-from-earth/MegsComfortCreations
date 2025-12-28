import { FieldErrors, useForm } from 'react-hook-form';
import { standardSchemaResolver } from '@hookform/resolvers/standard-schema';
import { collectorFormSchema } from './collectorFormSchema';
import type { CollectorFormData } from './collectorFormSchema';

// This hook manages the state and logic for a media collector form.
const defaultValues: CollectorFormData = {
  orderNumber: '',
  customerName: '',
  collectionData: {
    books: [],
    movies: [],
    videoGames: [],
    albums: [],
  },
};

export function useCollectorForm() {
  const form = useForm<CollectorFormData>({
    resolver: standardSchemaResolver(collectorFormSchema),
    defaultValues,
    mode: 'onSubmit',
    reValidateMode: 'onChange',
  });

  const onSubmit = (data: CollectorFormData) => {
    console.log(data);
  };

  const onError = (errors: FieldErrors<CollectorFormData>) => {
    console.error(errors);
  };
  return {
    form,
    onSubmit,
    onError,
  };
}
