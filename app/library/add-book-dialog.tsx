import { ChangeEvent, useId, useRef, useState } from 'react';
import CheckIcon from '@mui/icons-material/Check';
import CloseIcon from '@mui/icons-material/Close';
import SearchIcon from '@mui/icons-material/Search';
import { trpc } from 'lib/trpc/client';
import { readFileAsBase64 } from 'lib/helpers/read-file-as-base64';
import Dialog from '@/components/ui/dialog';
import Modal from '@/components/ui/modal';
import { Form } from '@/components/ui/form';
import {
  DefaultCoverRibbon,
  SetDefaultRibbonButton,
} from '@/components/shared/default-cover-ribbon';
import { SpineColorEditor } from '@/components/shared/spine-color-editor';
import { typeChipClassNames } from '@/mediacollector/collected-covers-block';
import { CoverImage, UploadTile } from '@/mediacollector/cover-tile';
import type { MediaItemForm } from '@/mediacollector/collector-form/media-item-form-schema';
import { updateAddBookCovers, type AddBookCoverAction } from './cover-edit-rules';
import { AddBookGenreChips } from './library-genres';
import LibraryItemTextField from './library-item-text-field';
import { useLibraryItemForm } from './use-library-item-form';

const EMPTY_COVER_SLOT_COUNT = 3;

const sectionHeadingClassName =
  'm-0 text-[13px] font-extrabold tracking-[1px] text-[#6b645d] uppercase';
const buttonClassName =
  'inline-flex h-11 cursor-pointer items-center gap-2 rounded-[10px] px-4.5 font-extrabold disabled:cursor-not-allowed';
const secondaryButtonClassName = `${buttonClassName} border border-[#ddd5cb] bg-white text-[#2b2826]`;
const primaryButtonClassName = `${buttonClassName} border-0 bg-[#2a4637] text-white disabled:bg-[#e3ddd5] disabled:text-[#5f5851]`;

function AddBookCoverTile({
  image,
  coverNumber,
  isFocused,
  onFocus,
  onToggleKeep,
  onSetDefault,
}: {
  image: MediaItemForm['images'][number];
  coverNumber: number;
  isFocused: boolean;
  onFocus: () => void;
  onToggleKeep: () => void;
  onSetDefault: () => void;
}) {
  const isKept = image.selected;
  return (
    <div className="relative">
      <button
        type="button"
        aria-pressed={isFocused}
        aria-label={`Cover ${coverNumber}${image.isDefault ? ', default' : ''}: show and edit its spine color`}
        onClick={onFocus}
        className={`relative flex aspect-[2/3] w-full cursor-pointer overflow-hidden rounded-md border-0 p-0 ${
          isFocused ? 'shadow-[0_0_0_2px_#ffffff,0_0_0_4.5px_#2b2826]' : ''
        } ${isKept ? '' : 'opacity-55'}`}
      >
        <span
          className="w-2 shrink-0"
          style={{ backgroundColor: image.spineColor }}
        />
        <span className="relative flex-1">
          <CoverImage url={image.url} alt="book cover" />
        </span>
      </button>
      {image.isDefault ? (
        <DefaultCoverRibbon variant="saved" className="left-2.5" />
      ) : null}
      {isFocused && isKept && !image.isDefault ? (
        <SetDefaultRibbonButton
          coverNumber={coverNumber}
          onClick={onSetDefault}
          className="left-2.5"
        />
      ) : null}
      <button
        type="button"
        role="checkbox"
        aria-checked={isKept}
        onClick={onToggleKeep}
        className={`absolute top-1.5 right-1.5 flex h-6 w-6 cursor-pointer items-center justify-center rounded-md border-2 border-white p-0 text-white ${
          isKept
            ? 'bg-[#3d3935] shadow-[0_1px_3px_rgba(0,0,0,0.3)]'
            : 'bg-white shadow-[inset_0_0_0_1.5px_#8a837c,0_1px_3px_rgba(0,0,0,0.3)]'
        }`}
      >
        {isKept ? <CheckIcon sx={{ fontSize: 16 }} /> : null}
      </button>
    </div>
  );
}

export default function AddBookDialog({ onClose }: { onClose: () => void }) {
  const headingId = useId();
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [hasSearched, setHasSearched] = useState(false);
  const [focusedIndex, setFocusedIndex] = useState<number | null>(null);
  const {
    form,
    formId,
    onSubmit,
    submitError,
    isDuplicateBookDialogOpen,
    closeDuplicateBookDialog,
    coverSearchBanner,
    isSearchingCovers,
    onSearchCovers,
  } = useLibraryItemForm({ onClose });
  const { mutateAsync: uploadCoverImage } =
    trpc.collect.uploadCoverImage.useMutation();

  const title = form.watch('blockInfo.title');
  const author = form.watch('blockInfo.author');
  const images = form.watch('images');
  const focusedImage =
    focusedIndex == null ? undefined : images[focusedIndex];
  const keptCount = images.filter((image) => image.selected).length;
  const isSubmitting = form.formState.isSubmitting;

  const updateCovers = (
    action: AddBookCoverAction,
    currentImages = form.getValues('images'),
  ) => {
    const next = updateAddBookCovers(
      { images: currentImages, focusedIndex },
      action,
    );
    form.setValue('images', next.images);
    const nextDefault = next.images.find((image) => image.isDefault);
    if (nextDefault) {
      form.setValue('blockInfo.spineColor', nextDefault.spineColor);
    }
    setFocusedIndex(next.focusedIndex);
  };

  const handleSearch = async () => {
    await onSearchCovers();
    if (
      !form.getFieldState('blockInfo.title').invalid &&
      !form.getFieldState('blockInfo.author').invalid
    ) {
      setHasSearched(true);
      const searchedImages = form.getValues('images');
      const defaultIndex = searchedImages.findIndex(
        (image) => image.isDefault,
      );
      setFocusedIndex(
        defaultIndex >= 0 ? defaultIndex : searchedImages.length > 0 ? 0 : null,
      );
    }
  };

  const handleUploadImage = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    setIsUploading(true);
    try {
      const currentImages = form.getValues('images');
      const fallbackSpineColor =
        currentImages.find((image) => image.isDefault)?.spineColor ??
        form.getValues('blockInfo.spineColor');
      const dataBase64 = await readFileAsBase64(file);
      const uploadedImage = await uploadCoverImage({
        blockID: form.getValues('blockID'),
        sortOrder: currentImages.length,
        fileName: file.name,
        mimeType: file.type,
        dataBase64,
      });
      const withUpload = [
        ...currentImages,
        {
          url: uploadedImage.url,
          selected: false,
          isDefault: false,
          spineColor: uploadedImage.palette[0] ?? fallbackSpineColor,
          uploaded: true,
          palette: uploadedImage.palette,
        },
      ];
      updateCovers(
        { type: 'toggleKeep', index: withUpload.length - 1 },
        withUpload,
      );
    } catch (error) {
      console.error('Failed to upload cover for new book', error);
    } finally {
      setIsUploading(false);
      event.target.value = '';
    }
  };

  return (
    <>
      <Modal aria-labelledby={headingId}>
        <Form {...form}>
          <form
            id={formId}
            onSubmit={form.handleSubmit(onSubmit)}
            className="flex max-h-[90vh] w-200 max-w-full flex-col overflow-hidden rounded-[20px] bg-white text-[#2b2826] shadow-[0_24px_60px_rgba(30,30,25,0.3)]"
          >
            <div className="flex items-center justify-between gap-3 border-b border-[#e3e9de] bg-[#f3f6f0] px-6 py-3.5">
              <div className="flex items-center gap-3">
                <span
                  className={`rounded-full px-2.5 py-1 text-xs font-extrabold ${typeChipClassNames.book}`}
                >
                  Book
                </span>
                <h2 id={headingId} className="m-0 text-[22px] font-extrabold">
                  Add book
                </h2>
              </div>
              <button
                type="button"
                aria-label="Close"
                onClick={onClose}
                className="flex h-11 w-11 shrink-0 cursor-pointer items-center justify-center rounded-[10px] border-0 bg-transparent text-[#6b645d]"
              >
                <CloseIcon />
              </button>
            </div>

            <div className="flex flex-col gap-4 overflow-y-auto px-1.5 pt-5 pb-2">
              <div className="mx-4.5 flex items-start gap-4.5">
                {focusedImage ? (
                  <div className="flex shrink-0 overflow-hidden rounded-r-md">
                    <div
                      className="w-2.5 shrink-0"
                      style={{ backgroundColor: focusedImage.spineColor }}
                    />
                    <div className="relative h-45 w-30">
                      <CoverImage
                        url={focusedImage.url}
                        alt={`${title} cover`}
                      />
                      {focusedImage.isDefault ? (
                        <DefaultCoverRibbon variant="saved" />
                      ) : null}
                    </div>
                  </div>
                ) : (
                  <div className="flex h-45 w-30 shrink-0 items-center justify-center rounded-md border-2 border-dashed border-[#d8d0c6] bg-[#fcfbf9] p-2 text-center text-[13px] font-extrabold text-[#6b645d]">
                    Search for covers
                  </div>
                )}
                <div className="flex min-w-0 flex-1 flex-col gap-2.5">
                  <LibraryItemTextField name="title" label="Title" />
                  <LibraryItemTextField name="author" label="Author" />
                  <div className="grid grid-cols-2 gap-2.5">
                    <LibraryItemTextField name="pubYear" label="Published" />
                    <LibraryItemTextField name="pageCount" label="Pages" />
                  </div>
                  <AddBookGenreChips />
                </div>
              </div>

              <div className="mx-4.5 flex items-center justify-between gap-3">
                <h3 className={sectionHeadingClassName}>
                  {hasSearched
                    ? `Covers · ${keptCount} of ${images.length} kept`
                    : 'Covers'}
                </h3>
                <button
                  type="button"
                  disabled={isSearchingCovers}
                  onClick={() => void handleSearch()}
                  className={
                    hasSearched
                      ? secondaryButtonClassName
                      : primaryButtonClassName
                  }
                >
                  <SearchIcon fontSize="small" />
                  {isSearchingCovers
                    ? 'Searching…'
                    : hasSearched
                      ? 'Search again'
                      : 'Search covers'}
                </button>
              </div>
              {coverSearchBanner ? (
                <p className="mx-4.5 my-0 text-xs font-bold text-[#9a3b26]">
                  {coverSearchBanner}
                </p>
              ) : null}

              <input
                ref={fileInputRef}
                type="file"
                accept="image/*"
                className="hidden"
                onChange={handleUploadImage}
              />
              <div
                className={`mx-4.5 grid grid-cols-5 gap-2.5 rounded-lg ${
                  hasSearched && keptCount === 0
                    ? 'outline-[1.5px] outline-offset-[6px] outline-[#b0533a] outline-dashed'
                    : ''
                }`}
              >
                {hasSearched ? (
                  <>
                    {images.map((image, imageIndex) => (
                      <AddBookCoverTile
                        key={`${image.url}-${imageIndex}`}
                        image={image}
                        coverNumber={imageIndex + 1}
                        isFocused={imageIndex === focusedIndex}
                        onFocus={() =>
                          updateCovers({ type: 'focus', index: imageIndex })
                        }
                        onToggleKeep={() =>
                          updateCovers({
                            type: 'toggleKeep',
                            index: imageIndex,
                          })
                        }
                        onSetDefault={() =>
                          updateCovers({
                            type: 'setDefault',
                            index: imageIndex,
                          })
                        }
                      />
                    ))}
                    <UploadTile
                      mediaType="book"
                      isUploading={isUploading}
                      onClick={() => fileInputRef.current?.click()}
                    />
                  </>
                ) : (
                  Array.from({ length: EMPTY_COVER_SLOT_COUNT }, (_, slot) => (
                    <div
                      key={slot}
                      className="aspect-[2/3] rounded-md border-[1.5px] border-dashed border-[#d8d0c6] bg-[#fcfbf9]"
                    />
                  ))
                )}
              </div>

              <div className={focusedImage ? '' : 'opacity-50'}>
                <SpineColorEditor
                  spineColor={focusedImage?.spineColor}
                  palette={focusedImage?.palette ?? []}
                  hasCover={focusedImage !== undefined}
                  isLocked={false}
                  isChanged={false}
                  caption={
                    focusedImage && focusedIndex != null
                      ? `Cover ${focusedIndex + 1} · ${focusedImage.spineColor}`
                      : 'Search for covers first'
                  }
                  onApply={(spineColor) =>
                    updateCovers({ type: 'setSpineColor', spineColor })
                  }
                />
              </div>

              {submitError ? (
                <p className="mx-4.5 my-0 text-xs font-bold text-red-600">
                  {submitError}
                </p>
              ) : null}
            </div>

            <div className="flex items-center justify-end gap-2.5 border-t border-[#ece6de] bg-[#f3f6f0] px-6 py-3.5">
              <button
                type="button"
                className={secondaryButtonClassName}
                onClick={onClose}
              >
                Cancel
              </button>
              <button
                type="submit"
                disabled={
                  title.trim() === '' ||
                  (author ?? '').trim() === '' ||
                  keptCount === 0 ||
                  isSubmitting
                }
                className={primaryButtonClassName}
              >
                {isSubmitting ? 'Adding…' : 'Add book'}
              </button>
            </div>
          </form>
        </Form>
      </Modal>
      {isDuplicateBookDialogOpen ? (
        <Dialog title="Book already exists" onClose={closeDuplicateBookDialog}>
          <p>A book with this title and author already exists.</p>
        </Dialog>
      ) : null}
    </>
  );
}
