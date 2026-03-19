import { useState } from 'react';
import Image from 'next/image';
import AddToPhotosIcon from '@mui/icons-material/AddToPhotos';
import { MediaType } from 'lib/constants/mediaTypes';

export interface MediaImageStripItem {
  url: string;
  selected?: boolean;
}

export interface MediaImageStripProps {
  mediaType: MediaType;
  images: MediaImageStripItem[];
  className?: string;
  albumTileClassName?: string;
  defaultTileClassName?: string;
  albumImageClassName?: string;
  defaultImageClassName?: string;
  showSelectionOverlay?: boolean;
  selectionOverlayText?: string;
  onImageClick?: (imageIndex: number) => void;
  showUploadButton?: boolean;
  isUploading?: boolean;
  onUploadButtonClick?: () => void;
  uploadButtonLabel?: string;
  uploadSlotLabel?: string;
}

const DEFAULT_ALBUM_TILE_CLASS =
  'relative z-10 h-31 w-31 overflow-hidden rounded-sm';
const DEFAULT_TILE_CLASS = 'relative z-10 h-31 w-21 overflow-hidden rounded-sm';
const DEFAULT_ALBUM_IMAGE_CLASS = 'cursor-pointer object-cover outline-2';
const DEFAULT_IMAGE_CLASS = 'cursor-pointer';

export default function MediaImageStrip({
  mediaType,
  images,
  className = 'flex flex-row items-center gap-3',
  albumTileClassName = DEFAULT_ALBUM_TILE_CLASS,
  defaultTileClassName = DEFAULT_TILE_CLASS,
  albumImageClassName = DEFAULT_ALBUM_IMAGE_CLASS,
  defaultImageClassName = DEFAULT_IMAGE_CLASS,
  showSelectionOverlay = false,
  selectionOverlayText = 'Selected',
  onImageClick,
  showUploadButton = false,
  isUploading = false,
  onUploadButtonClick,
  uploadButtonLabel = 'Add uploaded image',
  uploadSlotLabel = 'Uploading...',
}: MediaImageStripProps) {
  const [brokenImageUrls, setBrokenImageUrls] = useState<Record<string, boolean>>({});

  const tileClassName =
    mediaType === 'album' ? albumTileClassName : defaultTileClassName;
  const imageClassName =
    mediaType === 'album' ? albumImageClassName : defaultImageClassName;

  return (
    <div className={className}>
      {images.map((image, idx) => (
        <div
          className={tileClassName}
          key={`${image.url}-${idx}`}
          onClick={() => onImageClick?.(idx)}
        >
          {brokenImageUrls[image.url] ? (
            <div className="absolute inset-0 flex items-center justify-center rounded-sm bg-red-600/65 p-2 text-center text-sm font-bold text-white">
              Image path broken
            </div>
          ) : (
            <Image
              className={imageClassName}
              src={image.url}
              alt={`${mediaType} image`}
              fill
              sizes="(max-width: 640px) 33vw, (max-width: 1024px) 20vw, 200px"
              unoptimized
              loader={({ src }) => src}
              onError={() =>
                setBrokenImageUrls((previous) => ({
                  ...previous,
                  [image.url]: true,
                }))
              }
            />
          )}

          {showSelectionOverlay ? (
            <div
              className={`pointer-events-none absolute inset-0 flex content-center items-center ${
                image.selected ? 'opacity-100' : 'opacity-0'
              }`}
            >
              <p className='-translate-x-1 -rotate-65 font-["Just_Another_Hand"] text-5xl font-bold tracking-wider text-[rgb(0,77,0)]'>
                {selectionOverlayText}
              </p>
            </div>
          ) : null}
        </div>
      ))}

      {showUploadButton ? (
        <button
          type="button"
          className="flex h-31 w-21 cursor-pointer items-center justify-center rounded-sm border-2 border-dashed border-black/60 bg-white/45 transition-opacity hover:opacity-85 disabled:cursor-not-allowed disabled:opacity-50"
          onClick={onUploadButtonClick}
          disabled={isUploading}
          aria-label={uploadButtonLabel}
        >
          {isUploading ? (
            <span className="text-center text-sm font-bold">{uploadSlotLabel}</span>
          ) : (
            <AddToPhotosIcon sx={{ fontSize: '2.25rem' }} />
          )}
        </button>
      ) : null}
    </div>
  );
}
