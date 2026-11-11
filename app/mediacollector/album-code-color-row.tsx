import Image from 'next/image';
import ColorizeIcon from '@mui/icons-material/Colorize';
import { spotifyCodeUrl } from 'lib/trpc/routers/collect/actions/spotify-code';

const GHOST_BAR_HEIGHTS = [
  8, 11, 6.5, 12.5, 9.5, 14, 8, 11, 6.5, 9.5, 12.5, 5, 12.5, 9.5, 6.5, 11, 8,
  14, 9.5, 12.5, 6.5, 11, 8,
];

function MiniCode({
  spotifyAlbumUri,
  codeColor,
}: {
  spotifyAlbumUri: string | null | undefined;
  codeColor: string | null | undefined;
}) {
  const previewUrl =
    spotifyAlbumUri && codeColor
      ? spotifyCodeUrl(spotifyAlbumUri, codeColor)
      : null;

  if (previewUrl) {
    return (
      <span
        aria-hidden="true"
        className="relative h-5.5 w-13.5 shrink-0 overflow-hidden rounded-sm shadow-[inset_0_0_0_1px_rgba(0,0,0,0.12)]"
        style={{ backgroundColor: codeColor ?? undefined }}
      >
        <Image
          src={previewUrl}
          alt=""
          fill
          sizes="54px"
          unoptimized
          loader={({ src }) => src}
          className="object-contain"
        />
      </span>
    );
  }

  return (
    <span
      aria-hidden="true"
      className="box-border flex h-5.5 w-13.5 shrink-0 items-center justify-between overflow-hidden rounded-sm px-1 shadow-[inset_0_0_0_1px_rgba(0,0,0,0.12)]"
      style={{ backgroundColor: codeColor ?? '#efe9e2' }}
    >
      <span className="mr-1 h-3.5 w-3.5 shrink-0 rounded-full bg-[#1f1b18] opacity-35" />
      {GHOST_BAR_HEIGHTS.map((height, index) => (
        <span
          key={index}
          className="w-px shrink-0 rounded-full bg-[#1f1b18] opacity-35"
          style={{ height }}
        />
      ))}
    </span>
  );
}

export interface AlbumCodeColorRowProps {
  spotifyAlbumUri: string | null | undefined;
  codeColor: string | null | undefined;
  palette: string[];
  coverNumber: number | null;
  isLocked: boolean;
  isChanged: boolean;
  onApply: (hex: string) => void;
}

export default function AlbumCodeColorRow({
  spotifyAlbumUri,
  codeColor,
  palette,
  coverNumber,
  isLocked,
  isChanged,
  onApply,
}: AlbumCodeColorRowProps) {
  const hasCover = coverNumber !== null;
  const chosenColor = codeColor?.toLowerCase() ?? null;
  const caption = !hasCover
    ? 'Add a cover first'
    : isChanged
      ? 'Changed'
      : chosenColor
        ? `For cover ${coverNumber}`
        : `Not set · cover ${coverNumber}`;
  const isCaptionAmber = hasCover && (isChanged || chosenColor === null);

  const handleColorPick = async () => {
    if (!window.EyeDropper) {
      console.log('EyeDropper API not supported in this browser');
      return;
    }
    try {
      const { sRGBHex } = await new window.EyeDropper().open();
      onApply(sRGBHex.toLowerCase());
    } catch (error) {
      console.log(error);
    }
  };

  return (
    <div
      className={`mx-4.5 mb-3.5 flex flex-nowrap items-center gap-2.5 rounded-xl px-3 py-2.5 ${
        isChanged
          ? 'bg-[#fffaf0] shadow-[inset_0_0_0_1.5px_#c9973f]'
          : 'bg-[#f7f4f0]'
      }`}
    >
      <MiniCode spotifyAlbumUri={spotifyAlbumUri} codeColor={chosenColor} />
      <div className="flex min-w-0 flex-1 flex-col gap-px">
        <span className="text-xs font-extrabold">Code color</span>
        <span
          className={`truncate text-[11px] font-extrabold ${
            isCaptionAmber ? 'text-[#8a5a22]' : 'text-[#6b645d]'
          }`}
        >
          {caption}
        </span>
      </div>
      {hasCover && palette.length > 0 ? (
        <div
          role="radiogroup"
          aria-label="Code colors from cover"
          className="flex shrink-0 gap-1.5"
        >
          {palette.map((hex) => {
            const isChosen = hex.toLowerCase() === chosenColor;
            return (
              <button
                key={hex}
                type="button"
                role="radio"
                aria-checked={isChosen}
                aria-label={`Code color ${hex}`}
                disabled={isLocked}
                onClick={() => onApply(hex.toLowerCase())}
                className={`h-7 w-7 cursor-pointer rounded-full p-0 disabled:cursor-not-allowed ${
                  isChosen
                    ? 'border-0 shadow-[0_0_0_2px_#ffffff,0_0_0_4px_#2b2826]'
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
        aria-label="Pick code color with eyedropper"
        disabled={!hasCover || isLocked}
        onClick={() => void handleColorPick()}
        className="flex h-9 w-9 shrink-0 cursor-pointer items-center justify-center rounded-lg border border-[#e0d9d0] bg-white p-0 text-[#2b2826] disabled:cursor-not-allowed disabled:text-[#b3aaa0]"
      >
        <ColorizeIcon fontSize="small" />
      </button>
    </div>
  );
}
