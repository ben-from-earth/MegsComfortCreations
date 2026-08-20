import { useEffect, useId, useRef, useState } from 'react';
import { useForm } from 'react-hook-form';
import { standardSchemaResolver } from '@hookform/resolvers/standard-schema';
import { trpc } from 'lib/trpc/client';
import { useDatabasePageContext } from 'lib/context/DatabasePageContext';
import {
  convertMediaItemFormToDatabaseItem,
  mediaItemFormSchema,
  type MediaItemForm,
} from '@/mediacollector/collector-form/mediaItemFormSchema';
import {
  DATABASE_EDIT_FAILED_MESSAGE,
  GENRE_UPDATE_FAILED_MESSAGE,
  toDatabaseEditDisplayError,
} from './database-edit-error-display';

export function useMediaItemForm({
  item,
  onClose,
}: {
  item: MediaItemForm;
  onClose: () => void;
}) {
  const formId = useId();
  const [submitError, setSubmitError] = useState<string | null>(null);
  const form = useForm<MediaItemForm>({
    resolver: standardSchemaResolver(mediaItemFormSchema),
    defaultValues: item,
    mode: 'onSubmit',
    reValidateMode: 'onChange',
  });

  const itemRef = useRef(item);
  itemRef.current = item;
  const initialGenresRef = useRef([...item.blockInfo.genres]);

  useEffect(() => {
    const nextItem = itemRef.current;
    form.reset(nextItem);
    initialGenresRef.current = [...nextItem.blockInfo.genres];
  }, [form, item.blockID]);

  const { handleGetMedia } = useDatabasePageContext();
  const { mutateAsync: databaseEdit } = trpc.database.edit.useMutation();
  const { mutateAsync: linkGenres } = trpc.genres.link.useMutation();
  const { mutateAsync: unlinkGenres } = trpc.genres.unlink.useMutation();
  const utils = trpc.useUtils();

  const onSubmit = async (data: MediaItemForm) => {
    setSubmitError(null);
    form.clearErrors();

    let result;
    try {
      result = await databaseEdit({
        type: data.type,
        item: convertMediaItemFormToDatabaseItem(data),
      });
    } catch {
      setSubmitError(DATABASE_EDIT_FAILED_MESSAGE);
      return;
    }

    if (result.error != null) {
      const displayError = toDatabaseEditDisplayError(result.error);
      if (displayError.placement === 'field') {
        form.setError(displayError.field, {
          type: 'server',
          message: displayError.message,
        });
        return;
      }
      setSubmitError(displayError.message);
      return;
    }

    if (data.type === 'book') {
      const initialGenres = initialGenresRef.current;
      const nextGenres = data.blockInfo.genres;
      const genresToLink = nextGenres.filter(
        (genre) => !initialGenres.includes(genre),
      );
      const genresToUnlink = initialGenres.filter(
        (genre) => !nextGenres.includes(genre),
      );

      try {
        if (genresToLink.length > 0) {
          await linkGenres({ bookID: data.blockID, genres: genresToLink });
        }
        if (genresToUnlink.length > 0) {
          await unlinkGenres({ bookID: data.blockID, genres: genresToUnlink });
        }
      } catch {
        setSubmitError(GENRE_UPDATE_FAILED_MESSAGE);
        await handleGetMedia();
        return;
      }

      await utils.genres.getForBook.invalidate({ bookID: data.blockID });
    }

    await handleGetMedia();
    onClose();
  };

  return {
    form,
    formId,
    onSubmit,
    submitError,
  };
}
