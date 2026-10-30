import { useFormContext, useWatch } from 'react-hook-form';
import ColorizeIcon from '@mui/icons-material/Colorize';
import { CollectorFormData } from './collector-form/collector-form-schema';

export default function SpineColorRow({
  index,
  isDatabase,
}: {
  index: number;
  isDatabase: boolean;
}) {
  const { control, getValues, setValue } = useFormContext<CollectorFormData>();
  const spineColor = useWatch({
    control,
    name: `collectedData.${index}.blockInfo.spineColor`,
  });
  const images =
    useWatch({ control, name: `collectedData.${index}.images` }) ?? [];
  const hasSelectedImage = images.some((image) => image.selected);

  const handleColorPick = async () => {
    if (!window.EyeDropper) {
      console.log('EyeDropper API not supported in this browser');
      return;
    }
    const eyeDropper = new window.EyeDropper();
    try {
      const { sRGBHex } = await eyeDropper.open();
      const currentImages = getValues(`collectedData.${index}.images`);
      setValue(
        `collectedData.${index}.images`,
        currentImages.map((image) =>
          image.selected ? { ...image, spineColor: sRGBHex } : image,
        ),
      );
      setValue(`collectedData.${index}.blockInfo.spineColor`, sRGBHex);
    } catch (error) {
      console.log(error);
    }
  };

  const caption = !hasSelectedImage
    ? 'Pick a cover first'
    : isDatabase
      ? `Saved · ${spineColor}`
      : spineColor;

  return (
    <div className="mx-4.5 mb-3.5 flex flex-wrap items-center gap-2.5 rounded-xl bg-[#f7f4f0] px-3 py-2.5">
      <span
        className={`h-10 w-3 rounded-[3px] ${
          hasSelectedImage
            ? 'shadow-[inset_0_0_0_1px_rgba(0,0,0,0.15)]'
            : 'border-[1.5px] border-dashed border-[#b3aaa0]'
        }`}
        style={hasSelectedImage ? { backgroundColor: spineColor } : undefined}
      />
      <div className="flex min-w-23 flex-col gap-px">
        <span className="text-xs font-extrabold">Spine color</span>
        <span className="text-[11px] font-bold text-[#6b645d]">{caption}</span>
      </div>
      <button
        type="button"
        disabled={!hasSelectedImage}
        onClick={() => void handleColorPick()}
        className="ml-auto flex h-9 w-9 cursor-pointer items-center justify-center rounded-lg border border-[#e0d9d0] bg-white p-0 text-[#2b2826] disabled:cursor-not-allowed disabled:text-[#b3aaa0]"
      >
        <ColorizeIcon fontSize="small" />
      </button>
    </div>
  );
}
