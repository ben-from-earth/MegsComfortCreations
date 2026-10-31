import { useEffect, useId, useState } from 'react';
import CloseIcon from '@mui/icons-material/Close';
import { MediaType } from 'lib/constants/media-types';
import type { MediaItemForm } from './collector-form/collector-form-schema';
import { CoverTile, UploadTile } from './cover-tile';

const secondaryButtonClassName =
  'inline-flex h-11 cursor-pointer items-center rounded-[10px] border border-[#ddd5cb] bg-white px-4.5 font-extrabold text-[#2b2826]';
const primaryButtonClassName =
  'inline-flex h-11 cursor-pointer items-center rounded-[10px] border-0 bg-[#2a4637] px-5 font-extrabold text-white';

export default function CoversDialog({
  title,
  isDatabase,
  mediaType,
  images,
  isUploading,
  onCancel,
  onConfirm,
  onUploadClick,
}: {
  title: string;
  isDatabase: boolean;
  mediaType: MediaType;
  images: MediaItemForm['images'];
  isUploading: boolean;
  onCancel: () => void;
  onConfirm: (imageIndex: number) => void;
  onUploadClick?: () => void;
}) {
  const headingId = useId();
  const [pickedIndex, setPickedIndex] = useState(() =>
    Math.max(
      images.findIndex((image) => image.selected),
      0,
    ),
  );

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        onCancel();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [onCancel]);

  return (
    <div
      className="fixed inset-0 z-100 flex items-center justify-center bg-[rgba(43,48,40,0.55)] p-4"
      onClick={onCancel}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby={headingId}
        onClick={(event) => event.stopPropagation()}
        className="flex max-h-[90vh] w-190 max-w-full flex-col overflow-hidden rounded-[20px] bg-white text-[#2b2826] shadow-[0_24px_60px_rgba(30,30,25,0.3)]"
      >
        <div className="flex items-start justify-between gap-4 px-6 pt-5.5 pb-3.5">
          <div className="flex flex-col gap-1">
            <h2 id={headingId} className="m-0 text-[22px] font-extrabold">
              {title.trim() || 'Untitled'} ·{' '}
              {isDatabase ? 'saved covers' : 'covers'}
            </h2>
            <p className="m-0 text-[13px] font-semibold text-[#6b645d]">
              {images.length} covers{isDatabase ? ' in the database' : ''}.
              Pick the one to use for this order.
            </p>
          </div>
          <button
            type="button"
            onClick={onCancel}
            className="flex h-11 w-11 shrink-0 cursor-pointer items-center justify-center rounded-[10px] border-0 bg-transparent text-[#6b645d]"
          >
            <CloseIcon />
          </button>
        </div>

        <div
          role="radiogroup"
          className="grid grid-cols-[repeat(auto-fill,minmax(100px,1fr))] gap-3.5 overflow-y-auto px-6 pt-1.5 pb-5.5"
        >
          {images.map((image, imageIndex) => (
            <CoverTile
              key={`${image.url}-${imageIndex}`}
              url={image.url}
              mediaType={mediaType}
              isSelected={imageIndex === pickedIndex}
              isDefault={isDatabase && image.isDefault}
              onClick={() => setPickedIndex(imageIndex)}
            />
          ))}
          {onUploadClick ? (
            <UploadTile
              mediaType={mediaType}
              isUploading={isUploading}
              onClick={onUploadClick}
            />
          ) : null}
        </div>

        <div className="flex items-center justify-end gap-2.5 border-t border-[#ece6de] bg-[#f3f6f0] px-6 py-3.5">
          <button
            type="button"
            className={secondaryButtonClassName}
            onClick={onCancel}
          >
            Cancel
          </button>
          <button
            type="button"
            className={primaryButtonClassName}
            onClick={() => onConfirm(pickedIndex)}
          >
            Use this cover
          </button>
        </div>
      </div>
    </div>
  );
}
