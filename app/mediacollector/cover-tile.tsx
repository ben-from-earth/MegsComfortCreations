import { useState } from 'react';
import Image from 'next/image';
import AddIcon from '@mui/icons-material/Add';
import CheckIcon from '@mui/icons-material/Check';
import { MediaType } from 'lib/constants/media-types';
import {
  DefaultCoverRibbon,
  SetDefaultRibbonButton,
  defaultCoverTileLabel,
  type DefaultCoverRibbonVariant,
} from '@/components/shared/default-cover-ribbon';

export function coverAspectClassName(mediaType: MediaType) {
  return mediaType === 'album' ? 'aspect-square' : 'aspect-[2/3]';
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
  coverNumber,
  isSelected,
  defaultRibbon,
  onClick,
  onSetDefault,
}: {
  url: string;
  mediaType: MediaType;
  coverNumber: number;
  isSelected: boolean;
  defaultRibbon: DefaultCoverRibbonVariant | null;
  onClick: () => void;
  onSetDefault?: () => void;
}) {
  return (
    <div className="relative">
      <button
        type="button"
        role="radio"
        aria-checked={isSelected}
        aria-label={defaultCoverTileLabel(coverNumber, defaultRibbon)}
        onClick={onClick}
        className={`relative w-full cursor-pointer rounded-md p-0 ${coverAspectClassName(mediaType)} ${
          isSelected
            ? 'border-0 shadow-[0_0_0_2px_#ffffff,0_0_0_4.5px_#2a4637]'
            : 'border border-black/10'
        }`}
      >
        <span className="absolute inset-0 overflow-hidden rounded-md">
          <CoverImage url={url} alt={`${mediaType} cover`} />
        </span>
        {isSelected ? (
          <span className="absolute top-1.25 right-1.25 flex h-5 w-5 items-center justify-center rounded-full bg-white text-[#2b2826] shadow-[0_1px_3px_rgba(0,0,0,0.25)]">
            <CheckIcon sx={{ fontSize: 14 }} />
          </span>
        ) : null}
        {defaultRibbon ? <DefaultCoverRibbon variant={defaultRibbon} /> : null}
      </button>
      {onSetDefault ? (
        <SetDefaultRibbonButton
          coverNumber={coverNumber}
          onClick={onSetDefault}
        />
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
