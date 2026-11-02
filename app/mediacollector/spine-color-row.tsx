import { useFormContext, useWatch } from 'react-hook-form';
import { SpineColorEditor } from '@/components/shared/spine-color-editor';
import { CollectorFormData } from './collector-form/collector-form-schema';

export default function SpineColorRow({
  index,
  isLocked,
  isChanged,
}: {
  index: number;
  isLocked: boolean;
  isChanged: boolean;
}) {
  const { control, getValues, setValue } = useFormContext<CollectorFormData>();
  const spineColor = useWatch({
    control,
    name: `collectedData.${index}.blockInfo.spineColor`,
  });
  const images =
    useWatch({ control, name: `collectedData.${index}.images` }) ?? [];
  const selectedImage = images.find((image) => image.selected);

  const applySpineColor = (hex: string) => {
    const currentImages = getValues(`collectedData.${index}.images`);
    setValue(
      `collectedData.${index}.images`,
      currentImages.map((image) =>
        image.selected ? { ...image, spineColor: hex } : image,
      ),
    );
    setValue(`collectedData.${index}.blockInfo.spineColor`, hex);
  };

  return (
    <SpineColorEditor
      spineColor={spineColor}
      palette={selectedImage?.palette ?? []}
      hasCover={selectedImage !== undefined}
      isLocked={isLocked}
      isChanged={isChanged}
      onApply={applySpineColor}
    />
  );
}
