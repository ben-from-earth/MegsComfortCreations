import { useState } from 'react';
import Image from 'next/image';
import AddIcon from '@mui/icons-material/Add';
import CheckIcon from '@mui/icons-material/Check';
import { MediaType } from 'lib/constants/media-types';

export type CoverDefaultBand = 'default' | 'newDefault';

export function coverAspectClassName(mediaType: MediaType) {
  return mediaType === 'album' ? 'aspect-square' : 'aspect-[2/3]';
}

export function coverDefaultBand(
  image: { url: string; isDefault: boolean },
  originalDefaultUrl: string | undefined,
): CoverDefaultBand | null {
  if (!image.isDefault) {
    return null;
  }
  return originalDefaultUrl === undefined || image.url === originalDefaultUrl
    ? 'default'
    : 'newDefault';
}

export function CoverImage({ url, alt }: { url: string; alt: string }) {
  const [brokenUrl, setBrokenUrl] = useState<string | null>(null);

  if (brokenUrl === url) {
    return (
      <div className="absolute inset-0 flex items-center justify-center bg-[#f7e3dc] p-1 text-center text-[11px] font-bold text-[#9a3b26]">
        Image path broken
      </div>
    );
  }

  return (
    <Image
      className="object-cover"
      src={url}
      alt={alt}
      fill
      sizes="200px"
      unoptimized
      loader={({ src }) => src}
      onError={() => setBrokenUrl(url)}
    />
  );
}

export function CoverTile({
  url,
  mediaType,
  isSelected,
  defaultBand,
  onClick,
  onSetDefault,
}: {
  url: string;
  mediaType: MediaType;
  isSelected: boolean;
  defaultBand: CoverDefaultBand | null;
  onClick: () => void;
  onSetDefault?: () => void;
}) {
  return (
    <div className="relative">
      <button
        type="button"
        role="radio"
        aria-checked={isSelected}
        onClick={onClick}
        className={`relative w-full cursor-pointer overflow-hidden rounded-md p-0 ${coverAspectClassName(mediaType)} ${
          isSelected
            ? 'border-0 shadow-[0_0_0_2px_#ffffff,0_0_0_4.5px_#2a4637]'
            : 'border border-black/10'
        }`}
      >
        <CoverImage url={url} alt={`${mediaType} cover`} />
        {isSelected ? (
          <span className="absolute top-1.25 right-1.25 flex h-5 w-5 items-center justify-center rounded-full bg-white text-[#2b2826] shadow-[0_1px_3px_rgba(0,0,0,0.25)]">
            <CheckIcon sx={{ fontSize: 14 }} />
          </span>
        ) : null}
        {defaultBand ? (
          <span
            className={`absolute inset-x-0 bottom-0 py-1 text-center text-[10px] font-extrabold tracking-[0.5px] text-white uppercase ${
              defaultBand === 'newDefault'
                ? 'bg-[rgba(138,90,34,0.92)]'
                : 'bg-[rgba(43,40,38,0.82)]'
            }`}
          >
            {defaultBand === 'newDefault' ? 'New default' : 'Default'}
          </span>
        ) : null}
      </button>
      {onSetDefault ? (
        <button
          type="button"
          onClick={onSetDefault}
          className="absolute inset-x-0 bottom-0 cursor-pointer rounded-b-md border-0 bg-white/95 py-1 text-center text-[10px] font-extrabold tracking-[0.5px] text-[#2b2826] uppercase"
        >
          Set default
        </button>
      ) : null}
    </div>
  );
}

export function UploadTile({
  mediaType,
  isUploading,
  isDisabled = false,
  onClick,
}: {
  mediaType: MediaType;
  isUploading: boolean;
  isDisabled?: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={isUploading || isDisabled}
      className={`flex w-full cursor-pointer flex-col items-center justify-center gap-1 rounded-md border-[1.5px] border-dashed border-[#b3aaa0] bg-[#fcfbf9] p-0 text-[#6b645d] disabled:cursor-not-allowed disabled:opacity-60 ${coverAspectClassName(mediaType)}`}
    >
      {isUploading ? (
        <span className="text-[11px] font-extrabold">Uploading…</span>
      ) : (
        <>
          <AddIcon fontSize="small" />
          <span className="text-[11px] font-extrabold">Upload</span>
        </>
      )}
    </button>
  );
}
