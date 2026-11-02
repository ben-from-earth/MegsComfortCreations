import { useFormContext, useWatch } from 'react-hook-form';
import { SpineColorEditor } from '@/components/shared/spine-color-editor';
import type { MediaItemForm } from '@/mediacollector/collector-form/media-item-form-schema';

export default function LibrarySpineColorRow({
  coverRole,
  isLocked,
  isChanged,
}: {
  coverRole: 'selected' | 'isDefault';
  isLocked: boolean;
  isChanged: boolean;
}) {
  const { control, getValues, setValue } = useFormContext<MediaItemForm>();
  const images = useWatch({ control, name: 'images' }) ?? [];
  const cover = images.find((image) => image[coverRole]);

  const applySpineColor = (hex: string) => {
    setValue(
      'images',
      getValues('images').map((image) =>
        image[coverRole] ? { ...image, spineColor: hex } : image,
      ),
    );
    setValue('blockInfo.spineColor', hex);
  };

  return (
    <SpineColorEditor
      spineColor={cover?.spineColor}
      palette={cover?.palette ?? []}
      hasCover={cover !== undefined}
      isLocked={isLocked}
      isChanged={isChanged}
      onApply={applySpineColor}
    />
  );
}
