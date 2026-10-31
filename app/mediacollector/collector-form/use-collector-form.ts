import { useId } from 'react';
import { useForm } from 'react-hook-form';
import { standardSchemaResolver } from '@hookform/resolvers/standard-schema';
import { collectorFormSchema } from './collector-form-schema';
import type { CollectorFormData } from './collector-form-schema';
import { convertMediaItemFormToDatabaseItem } from './media-item-form-schema';
import type { DatabaseItemEdit } from '../database-item-changes';
import type {
  DatabaseSaveResultItem,
  DatabaseSaveServerResponse,
} from 'lib/interfaces/global-interfaces';
import { trpc } from 'lib/trpc/client';

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
  pngExportBlockIds: [],
  pngFormat: null,
};

export function useCollectorForm() {
  const formId = useId();
  const form = useForm<CollectorFormData>({
    resolver: standardSchemaResolver(collectorFormSchema),
    defaultValues,
    mode: 'onSubmit',
    reValidateMode: 'onChange',
  });

  const { mutateAsync: databaseSave } = trpc.database.save.useMutation();
  const { mutateAsync: databaseEdit } = trpc.database.edit.useMutation();
  const { mutateAsync: linkGenres } = trpc.genres.link.useMutation();
  const { mutateAsync: unlinkGenres } = trpc.genres.unlink.useMutation();

  const saveDatabaseEdit = async ({
    block,
    original,
  }: DatabaseItemEdit): Promise<DatabaseSaveResultItem> => {
    const title = block.blockInfo.title;
    const failure = (error: string, message: string, errors: string[]) => ({
      success: false as const,
      blockID: block.blockID,
      title,
      error,
      message,
      errors,
    });

    try {
      const result = await databaseEdit({
        type: block.type,
        item: convertMediaItemFormToDatabaseItem(block),
      });
      if (result.error != null) {
        return failure(result.error, result.message, result.errors);
      }
    } catch (error) {
      return failure(
        'Database Edit Error',
        'An error occurred while trying to save the edit',
        [error instanceof Error ? error.message : 'Unknown error'],
      );
    }

    if (block.type === 'book') {
      const genresToLink = block.blockInfo.genres.filter(
        (genre) => !original.blockInfo.genres.includes(genre),
      );
      const genresToUnlink = original.blockInfo.genres.filter(
        (genre) => !block.blockInfo.genres.includes(genre),
      );
      try {
        if (genresToLink.length > 0) {
          await linkGenres({ bookID: block.blockID, genres: genresToLink });
        }
        if (genresToUnlink.length > 0) {
          await unlinkGenres({ bookID: block.blockID, genres: genresToUnlink });
        }
      } catch (error) {
        return failure(
          'Genre Update Error',
          'The genres for this item could not be updated',
          [error instanceof Error ? error.message : 'Unknown error'],
        );
      }
    }

    return {
      success: true,
      blockID: block.blockID,
      title,
      message: `${title} successfully edited.`,
      type: block.type,
      actionAttemptItem: {
        id: block.blockID,
        title,
        spineColor: block.blockInfo.spineColor,
        images: block.images,
      },
    };
  };

  const onSubmit = async (
    data: CollectorFormData,
    databaseEdits: DatabaseItemEdit[],
  ): Promise<DatabaseSaveServerResponse> => {
    const editedBlockIds = new Set(
      databaseEdits.map(({ block }) => block.blockID),
    );
    const [saveResult, editResults] = await Promise.all([
      databaseSave(
        data.collectedData.filter(
          (block) => !editedBlockIds.has(block.blockID),
        ),
      ),
      Promise.all(databaseEdits.map(saveDatabaseEdit)),
    ]);
    return [...saveResult, ...editResults];
  };

  return {
    form,
    formId,
    onSubmit,
  };
}
