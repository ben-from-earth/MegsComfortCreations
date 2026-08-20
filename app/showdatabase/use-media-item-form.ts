import { useEffect, useId, useRef } from 'react';
import { type FieldErrors, useForm } from 'react-hook-form';
import { standardSchemaResolver } from '@hookform/resolvers/standard-schema';
import { trpc } from 'lib/trpc/client';
import { useDatabasePageContext } from 'lib/context/DatabasePageContext';
import {
  convertMediaItemFormToDatabaseItem,
  mediaItemFormSchema,
  type MediaItemForm,
} from '@/mediacollector/collector-form/mediaItemFormSchema';

export function useMediaItemForm({
  item,
  onClose,
}: {
  item: MediaItemForm;
  onClose: () => void;
}) {
  const formId = useId();
  const form = useForm<MediaItemForm>({
    resolver: standardSchemaResolver(mediaItemFormSchema),
    defaultValues: item,
    mode: 'onSubmit',
    reValidateMode: 'onChange',
  });

  const itemRef = useRef(item);
  itemRef.current = item;

  useEffect(() => {
    form.reset(itemRef.current);
  }, [form, item.blockID]);

  const { handleGetMedia } = useDatabasePageContext();
  const { mutateAsync: databaseEdit } = trpc.database.edit.useMutation();
  const { mutateAsync: linkGenres } = trpc.genres.link.useMutation();
  const { mutateAsync: unlinkGenres } = trpc.genres.unlink.useMutation();
  const utils = trpc.useUtils();
  const initialGenres = item.blockInfo.genres;

  const onSubmit = async (data: MediaItemForm) => {
    const result = await databaseEdit({
      type: data.type,
      item: convertMediaItemFormToDatabaseItem(data),
    });
    if ('error' in result) {
      onClose();
      return;
    }

    if (data.type === 'book') {
      const nextGenres = data.blockInfo.genres;
      const genresToLink = nextGenres.filter(
        (genre) => !initialGenres.includes(genre),
      );
      const genresToUnlink = initialGenres.filter(
        (genre) => !nextGenres.includes(genre),
      );

      if (genresToLink.length > 0) {
        try {
          await linkGenres({ bookID: data.blockID, genres: genresToLink });
        } catch {
          console.log('Genre link error');
        }
      }
      if (genresToUnlink.length > 0) {
        try {
          await unlinkGenres({ bookID: data.blockID, genres: genresToUnlink });
        } catch {
          console.log('Genre unlink error');
        }
      }

      await utils.genres.getForBook.invalidate({ bookID: data.blockID });
    }

    await handleGetMedia();
    onClose();
  };

  const onError = (errors: FieldErrors<MediaItemForm>) => {
    console.error(errors);
  };

  return {
    form,
    formId,
    onSubmit,
    onError,
  };
}
