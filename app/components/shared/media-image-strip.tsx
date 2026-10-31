import { useState } from 'react';
import Image from 'next/image';
import AddToPhotosIcon from '@mui/icons-material/AddToPhotos';
import Popover from '@mui/material/Popover';
import { MediaType } from 'lib/constants/media-types';
import {
  DefaultCoverRibbon,
  SetDefaultRibbonButton,
  defaultCoverRibbonVariant,
  defaultCoverTileLabel,
} from '@/components/shared/default-cover-ribbon';

export interface MediaImageStripItem {
  url: string;
  selected?: boolean;
  isDefault?: boolean;
  spineColor?: string;
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
  originalDefaultUrl?: string;
  onSetDefault?: (imageIndex: number) => void;
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
  originalDefaultUrl,
  onSetDefault,
  showUploadButton = false,
  isUploading = false,
  onUploadButtonClick,
  uploadButtonLabel = 'Add uploaded image',
  uploadSlotLabel = 'Uploading...',
}: MediaImageStripProps) {
  const [brokenImageUrls, setBrokenImageUrls] = useState<
    Record<string, boolean>
  >({});
  const [overflowAnchorElement, setOverflowAnchorElement] =
    useState<HTMLButtonElement | null>(null);

  const tileClassName =
    mediaType === 'album' ? albumTileClassName : defaultTileClassName;
  const imageClassName =
    mediaType === 'album' ? albumImageClassName : defaultImageClassName;
  const defaultImageIndex = Math.max(
    images.findIndex((image) => image.isDefault),
    0,
  );
  const selectedImageIndex = images.findIndex((image) => image.selected);
  const primaryImageIndex =
    selectedImageIndex >= 0 ? selectedImageIndex : defaultImageIndex;
  const primaryImage = images[primaryImageIndex];
  const secondaryImages = images
    .map((image, index) => ({ image, originalIndex: index }))
    .filter((record) => record.originalIndex !== primaryImageIndex);
  const secondaryVisibleImages = secondaryImages.slice(0, 3);
  const overflowImages = secondaryImages.slice(3);
  const showOverflowButton = overflowImages.length > 0;

  const tileImageClassName = `cursor-pointer ${imageClassName}`.trim();
  const ribbonVariant = (image: MediaImageStripItem) =>
    defaultCoverRibbonVariant(image, originalDefaultUrl);
  const primaryRibbon = primaryImage ? ribbonVariant(primaryImage) : null;
  const renderRibbon = (image: MediaImageStripItem) => {
    const variant = ribbonVariant(image);
    return variant ? <DefaultCoverRibbon variant={variant} /> : null;
  };

  const renderImageTile = (
    image: MediaImageStripItem,
    imageIndex: number,
    useGridImageClass = false,
  ) => (
    <div
      className={tileClassName}
      onClick={() => onImageClick?.(imageIndex)}
      role="button"
      tabIndex={0}
      aria-label={defaultCoverTileLabel(imageIndex + 1, ribbonVariant(image))}
      onKeyDown={(event) => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          onImageClick?.(imageIndex);
        }
      }}
    >
      {brokenImageUrls[image.url] ? (
        <div className="absolute inset-0 flex items-center justify-center rounded-sm bg-red-600/65 p-2 text-center text-sm font-bold text-white">
          Image path broken
        </div>
      ) : (
        <Image
          className={useGridImageClass ? 'object-cover' : tileImageClassName}
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
  );

  return (
    <div className={className}>
      {primaryImage ? (
        <div className="relative shrink-0">
          {renderImageTile(primaryImage, primaryImageIndex)}
          {primaryRibbon ? (
            <span className="pointer-events-none absolute inset-0 z-20">
              <DefaultCoverRibbon variant={primaryRibbon} />
            </span>
          ) : onSetDefault ? (
            <SetDefaultRibbonButton
              coverNumber={primaryImageIndex + 1}
              onClick={() => onSetDefault(primaryImageIndex)}
            />
          ) : null}
        </div>
      ) : null}
      {secondaryImages.length > 0 ? (
        <div className="grid h-31 w-21 grid-cols-2 grid-rows-2 gap-1 overflow-visible">
          {secondaryVisibleImages.map(({ image, originalIndex }) => (
            <button
              type="button"
              className="relative overflow-visible rounded-sm"
              key={`${image.url}-${originalIndex}`}
              onClick={() => onImageClick?.(originalIndex)}
              aria-label={defaultCoverTileLabel(
                originalIndex + 1,
                ribbonVariant(image),
              )}
            >
              {brokenImageUrls[image.url] ? (
                <div className="absolute inset-0 flex items-center justify-center rounded-sm bg-red-600/65 p-1 text-center text-[10px] font-bold text-white">
                  Broken
                </div>
              ) : (
                <Image
                  className="object-cover"
                  src={image.url}
                  alt={`${mediaType} image`}
                  fill
                  sizes="80px"
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
              {renderRibbon(image)}

            </button>
          ))}
          {showOverflowButton ? (
            <button
              type="button"
              className="rounded-sm border border-black/40 bg-black/70 text-xs font-bold text-white"
              onClick={(event) => setOverflowAnchorElement(event.currentTarget)}
            >
              +{overflowImages.length}
            </button>
          ) : null}
        </div>
      ) : null}
      <Popover
        open={Boolean(overflowAnchorElement)}
        anchorEl={overflowAnchorElement}
        onClose={() => setOverflowAnchorElement(null)}
        anchorOrigin={{ vertical: 'bottom', horizontal: 'left' }}
      >
        <div className="grid max-w-sm grid-cols-2 gap-2 p-2">
          {overflowImages.map(({ image, originalIndex }) => (
            <button
              type="button"
              key={`${image.url}-${originalIndex}-popover`}
              className="relative h-24 w-16 rounded-sm"
              aria-label={defaultCoverTileLabel(
                originalIndex + 1,
                ribbonVariant(image),
              )}
              onClick={() => {
                onImageClick?.(originalIndex);
                setOverflowAnchorElement(null);
              }}
            >
              <Image
                className="rounded-sm object-cover"
                src={image.url}
                alt={`${mediaType} image`}
                fill
                sizes="100px"
                unoptimized
                loader={({ src }) => src}
              />
              {renderRibbon(image)}

            </button>
          ))}
        </div>
      </Popover>

      {showUploadButton ? (
        <button
          type="button"
          className="flex h-31 w-21 cursor-pointer items-center justify-center rounded-sm border-2 border-dashed border-black/60 bg-white/45 transition-opacity hover:opacity-85 disabled:cursor-not-allowed disabled:opacity-50"
          onClick={onUploadButtonClick}
          disabled={isUploading}
          aria-label={uploadButtonLabel}
        >
          {isUploading ? (
            <span className="text-center text-sm font-bold">
              {uploadSlotLabel}
            </span>
          ) : (
            <AddToPhotosIcon sx={{ fontSize: '2.25rem' }} />
          )}
        </button>
      ) : null}
    </div>
  );
}
