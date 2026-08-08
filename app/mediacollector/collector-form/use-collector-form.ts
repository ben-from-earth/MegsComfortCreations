import { FieldErrors, useForm } from 'react-hook-form';
import { standardSchemaResolver } from '@hookform/resolvers/standard-schema';
import { collectorFormSchema } from './collectorFormSchema';
import type { CollectorFormData } from './collectorFormSchema';

// library imports
import { trpc } from 'lib/trpc/client';

// This hook manages the state and logic for a media collector form.
const defaultValues: CollectorFormData = {
  orderNumber: '',
  customerName: '',
  bookClubRepeat: 1,
  collectionList: {
    book: [],
    movie: [],
    videoGame: [],
    album: [],
  },
  collectedData: [],
  pngFormat: null,
};

export function useCollectorForm() {
  const form = useForm<CollectorFormData>({
    resolver: standardSchemaResolver(collectorFormSchema),
    defaultValues,
    mode: 'onSubmit',
    reValidateMode: 'onChange',
  });

  const { mutateAsync: databaseSave } = trpc.database.save.useMutation();

  const onSubmit = async (data: CollectorFormData) => {
    const saveResult = await databaseSave(data.collectedData);
    return saveResult;
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
