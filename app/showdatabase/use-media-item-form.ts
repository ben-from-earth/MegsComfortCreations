import { useEffect, useId, useRef, useState } from 'react';
import { useForm } from 'react-hook-form';
import { standardSchemaResolver } from '@hookform/resolvers/standard-schema';
import { trpc } from 'lib/trpc/client';
import { useDatabasePageContext } from 'lib/context/DatabasePageContext';
import {
  convertMediaItemFormToDatabaseItem,
  getMediaItemFormDefaultValues,
  mediaItemFormSchema,
  type MediaItemForm,
} from '@/mediacollector/collector-form/mediaItemFormSchema';
import {
  DATABASE_EDIT_FAILED_MESSAGE,
  GENRE_UPDATE_FAILED_MESSAGE,
  toDatabaseEditDisplayError,
} from './database-edit-error-display';

const COVER_SEARCH_DUPLICATE_MESSAGE =
  'A book with this title and author already exists.';
const COVER_SEARCH_FAILED_MESSAGE = 'Cover search failed. Please try again.';

export function useMediaItemForm({
  item,
  onClose,
}: {
  item?: MediaItemForm;
  onClose: () => void;
}) {
  const formId = useId();
  const isAdd = item == null;
  const [submitError, setSubmitError] = useState<string | null>(null);
  const [isDuplicateBookDialogOpen, setIsDuplicateBookDialogOpen] =
    useState(false);
  const [coverSearchBanner, setCoverSearchBanner] = useState<string | null>(
    null,
  );
  const [isSearchingCovers, setIsSearchingCovers] = useState(false);
  const defaultValues = getMediaItemFormDefaultValues(item);
  const form = useForm<MediaItemForm>({
    resolver: standardSchemaResolver(mediaItemFormSchema),
    defaultValues,
    mode: 'onSubmit',
    reValidateMode: 'onChange',
  });

  const itemRef = useRef(item);
  itemRef.current = item;
  const initialGenresRef = useRef([...defaultValues.blockInfo.genres]);
  const itemBlockID = item?.blockID;

  useEffect(() => {
    if (itemBlockID == null) {
      return;
    }
    const nextItem = itemRef.current;
    if (nextItem == null) {
      return;
    }
    form.reset(nextItem);
    initialGenresRef.current = [...nextItem.blockInfo.genres];
  }, [form, itemBlockID]);

  const { handleGetMedia } = useDatabasePageContext();
  const { mutateAsync: databaseEdit } = trpc.database.edit.useMutation();
  const { mutateAsync: databaseSave } = trpc.database.save.useMutation();
  const { mutateAsync: collectMedia } = trpc.collect.collectMedia.useMutation();
  const { mutateAsync: linkGenres } = trpc.genres.link.useMutation();
  const { mutateAsync: unlinkGenres } = trpc.genres.unlink.useMutation();
  const utils = trpc.useUtils();

  const applySaveDisplayError = (error: string | undefined) => {
    const displayError = toDatabaseEditDisplayError(error ?? 'Unknown');
    if (displayError.placement === 'field') {
      form.setError(displayError.field, {
        type: 'server',
        message: displayError.message,
      });
      return;
    }
    setSubmitError(displayError.message);
  };

  const onSubmit = async (data: MediaItemForm) => {
    setSubmitError(null);
    form.clearErrors();
    setIsDuplicateBookDialogOpen(false);

    if (isAdd) {
      let results;
      try {
        results = await databaseSave([data]);
      } catch {
        setSubmitError(DATABASE_EDIT_FAILED_MESSAGE);
        return;
      }

      const result = results[0];
      if (result == null || !result.success) {
        if (result?.error === 'Duplicate Book') {
          setIsDuplicateBookDialogOpen(true);
          return;
        }
        applySaveDisplayError(result?.error);
        return;
      }

      await handleGetMedia();
      onClose();
      return;
    }

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
      applySaveDisplayError(result.error);
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

  const searchCovers = async (data: MediaItemForm) => {
    setIsSearchingCovers(true);
    try {
      const blocks = await collectMedia({
        book: [
          {
            title: data.blockInfo.title,
            author: data.blockInfo.author ?? undefined,
          },
        ],
        movie: [],
        videoGame: [],
        album: [],
      });

      if (blocks.some((block) => block.isDatabase)) {
        setCoverSearchBanner(COVER_SEARCH_DUPLICATE_MESSAGE);
        return;
      }

      const newBlock = blocks[0];
      if (newBlock == null) {
        setCoverSearchBanner(COVER_SEARCH_FAILED_MESSAGE);
        return;
      }

      const currentImages = form.getValues('images');
      const appendedImages = newBlock.images.map((image, index) => ({
        ...image,
        isDefault: currentImages.length === 0 && index === 0,
      }));
      form.setValue(
        'images',
        currentImages.length === 0
          ? appendedImages
          : [...currentImages, ...appendedImages],
      );

      const currentBlockInfo = form.getValues('blockInfo');
      if (
        currentBlockInfo.pubYear == null &&
        newBlock.blockInfo.pubYear != null
      ) {
        form.setValue('blockInfo.pubYear', newBlock.blockInfo.pubYear);
      }
      if (
        currentBlockInfo.pageCount == null &&
        newBlock.blockInfo.pageCount != null
      ) {
        form.setValue('blockInfo.pageCount', newBlock.blockInfo.pageCount);
      }
    } catch {
      setCoverSearchBanner(COVER_SEARCH_FAILED_MESSAGE);
    } finally {
      setIsSearchingCovers(false);
    }
  };

  const onSearchCovers = async () => {
    setCoverSearchBanner(null);
    const isTitleValid = await form.trigger('blockInfo.title');
    if (!isTitleValid) {
      return;
    }
    await searchCovers(form.getValues());
  };

  return {
    form,
    formId,
    isAdd,
    onSubmit,
    submitError,
    isDuplicateBookDialogOpen,
    closeDuplicateBookDialog: () => setIsDuplicateBookDialogOpen(false),
    coverSearchBanner,
    isSearchingCovers,
    onSearchCovers,
  };
}
