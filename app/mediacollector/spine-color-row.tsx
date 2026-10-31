import { useFormContext, useWatch } from 'react-hook-form';
import ColorizeIcon from '@mui/icons-material/Colorize';
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
  const hasSelectedImage = selectedImage !== undefined;
  const palette = selectedImage?.palette ?? [];

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

  const handleColorPick = async () => {
    if (!window.EyeDropper) {
      console.log('EyeDropper API not supported in this browser');
      return;
    }
    const eyeDropper = new window.EyeDropper();
    try {
      const { sRGBHex } = await eyeDropper.open();
      applySpineColor(sRGBHex);
    } catch (error) {
      console.log(error);
    }
  };

  const normalizedSpineColor = spineColor?.toLowerCase();

  return (
    <div className="mx-4.5 mb-3.5 flex flex-wrap items-center gap-2.5 rounded-xl bg-[#f7f4f0] px-3 py-2.5">
      <span
        className={`h-10 w-3 rounded-[3px] ${
          hasSelectedImage
            ? 'shadow-[inset_0_0_0_1px_rgba(0,0,0,0.15)]'
            : 'border-[1.5px] border-dashed border-[#b3aaa0]'
        } ${isChanged ? 'outline-2 outline-offset-2 outline-[#c9973f]' : ''}`}
        style={hasSelectedImage ? { backgroundColor: spineColor } : undefined}
      />
      <div className="flex min-w-23 flex-col gap-px">
        <span className="text-xs font-extrabold">Spine color</span>
        {hasSelectedImage ? (
          <span
            className={`text-[11px] font-bold ${
              isChanged ? 'text-[#8a5a22]' : 'text-[#6b645d]'
            }`}
          >
            {spineColor}
          </span>
        ) : null}
      </div>
      {hasSelectedImage && palette.length > 0 ? (
        <div role="radiogroup" className="flex gap-1.5">
          {palette.map((hex) => {
            const isChosen = hex === normalizedSpineColor;
            return (
              <button
                key={hex}
                type="button"
                role="radio"
                aria-checked={isChosen}
                title={hex}
                disabled={isLocked}
                onClick={() => applySpineColor(hex)}
                className={`h-7 w-7 cursor-pointer rounded-full p-0 disabled:cursor-not-allowed ${
                  isChosen
                    ? 'border-2 border-white shadow-[0_0_0_2px_#2a4637]'
                    : 'border border-[rgba(0,0,0,0.15)]'
                } ${isLocked && !isChosen ? 'opacity-35' : ''}`}
                style={{ backgroundColor: hex }}
              />
            );
          })}
        </div>
      ) : null}
      <button
        type="button"
        disabled={!hasSelectedImage || isLocked}
        onClick={() => void handleColorPick()}
        className="ml-auto flex h-9 w-9 cursor-pointer items-center justify-center rounded-lg border border-[#e0d9d0] bg-white p-0 text-[#2b2826] disabled:cursor-not-allowed disabled:text-[#b3aaa0]"
      >
        <ColorizeIcon fontSize="small" />
      </button>
    </div>
  );
}
