'use client';

import { use, useEffect, useRef, useState } from 'react';
import type { RefObject } from 'react';
import CheckIcon from '@mui/icons-material/Check';
import CloseIcon from '@mui/icons-material/Close';
import GenreContext from 'lib/context/genre-context';
import {
  Command,
  CommandEmpty,
  CommandInput,
  CommandItem,
  CommandList,
} from '@/components/ui/command';
import { cn } from '@/lib/utils/classnames';

interface GenreEditActions {
  onAdd: (genre: string) => void;
  onRemove: (genre: string) => void;
}

type GenrePopover = 'add' | 'all';

const VISIBLE_GENRE_COUNT = 2;

const labelClassName =
  'text-[11px] font-extrabold tracking-[1px] text-[#6b645d] uppercase';
const popoverClassName =
  'absolute z-20 rounded-xl bg-white shadow-[0_12px_28px_rgba(40,35,30,0.18)]';
const addGenreButtonClassName =
  'h-7 shrink-0 cursor-pointer rounded-full border border-dashed border-[#b3aaa0] bg-white px-2.5 text-[11px] font-bold text-[#6b645d]';

function usePopoverDismiss(
  containerRef: RefObject<HTMLElement | null>,
  isOpen: boolean,
  onDismiss: (shouldReturnFocus: boolean) => void,
) {
  const onDismissRef = useRef(onDismiss);
  useEffect(() => {
    onDismissRef.current = onDismiss;
  });

  useEffect(() => {
    if (!isOpen) {
      return;
    }
    const handleOutsideInteraction = (event: Event) => {
      if (
        event.target instanceof Node &&
        !containerRef.current?.contains(event.target)
      ) {
        onDismissRef.current(false);
      }
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        onDismissRef.current(true);
      }
    };
    document.addEventListener('pointerdown', handleOutsideInteraction);
    document.addEventListener('focusin', handleOutsideInteraction);
    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('pointerdown', handleOutsideInteraction);
      document.removeEventListener('focusin', handleOutsideInteraction);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [isOpen, containerRef]);
}

function GenreTag({
  genre,
  onRemove,
  className,
}: {
  genre: string;
  onRemove?: () => void;
  className?: string;
}) {
  return (
    <span
      className={cn(
        'inline-flex h-7 max-w-full min-w-0 items-center gap-1 rounded-full bg-[#3d3935] text-[11px] font-bold text-white',
        onRemove ? 'pr-1 pl-2.5' : 'px-2.5',
        className,
      )}
    >
      <span className="truncate">{genre}</span>
      {onRemove ? (
        <button
          type="button"
          aria-label={`Remove ${genre}`}
          onClick={onRemove}
          className="flex h-5 w-5 shrink-0 cursor-pointer items-center justify-center rounded-full border-0 bg-transparent p-0 text-[#e6e0d8]"
        >
          <CloseIcon sx={{ fontSize: 14 }} />
        </button>
      ) : null}
    </span>
  );
}

function GenreSearchList({
  genres,
  onAdd,
  onRemove,
}: GenreEditActions & { genres: string[] }) {
  const allGenres = use(GenreContext);
  const [search, setSearch] = useState('');
  const genreSearch = search.trim().toLowerCase();
  const options = allGenres.filter((genre) =>
    genre.toLowerCase().includes(genreSearch),
  );

  return (
    <Command shouldFilter={false} className="bg-transparent">
      <CommandInput
        autoFocus
        value={search}
        onValueChange={setSearch}
        placeholder="Search genres…"
      />
      <CommandList className="max-h-42">
        <CommandEmpty className="p-3 text-xs text-[#6b645d] italic">
          No matching genres
        </CommandEmpty>
        {options.map((genre) => {
          const isChecked = genres.includes(genre);
          return (
            <CommandItem
              key={genre}
              value={genre}
              aria-checked={isChecked}
              onSelect={() => (isChecked ? onRemove(genre) : onAdd(genre))}
            >
              <span className="flex w-4 shrink-0 justify-center text-[#2a4637]">
                {isChecked ? <CheckIcon sx={{ fontSize: 16 }} /> : null}
              </span>
              {genre}
            </CommandItem>
          );
        })}
      </CommandList>
    </Command>
  );
}

function AddGenrePopover({
  className,
  ...searchListProps
}: GenreEditActions & { genres: string[]; className: string }) {
  return (
    <div className={cn(popoverClassName, 'w-55 overflow-hidden', className)}>
      <GenreSearchList {...searchListProps} />
    </div>
  );
}

function AllGenresPopover({
  genres,
  isEditable,
  onAdd,
  onRemove,
  onClose,
}: GenreEditActions & {
  genres: string[];
  isEditable: boolean;
  onClose: () => void;
}) {
  return (
    <div
      className={`${popoverClassName} top-11 right-1.5 left-1.5 flex flex-col gap-2.5 p-3`}
    >
      <div className="flex items-center justify-between gap-2">
        <span className={labelClassName}>All genres · {genres.length}</span>
        <button
          type="button"
          aria-label="Close"
          autoFocus={!isEditable}
          onClick={onClose}
          className="flex h-7 w-7 shrink-0 cursor-pointer items-center justify-center rounded-lg border-0 bg-transparent p-0 text-[#6b645d]"
        >
          <CloseIcon sx={{ fontSize: 18 }} />
        </button>
      </div>
      <div className="flex flex-wrap gap-1.5">
        {genres.map((genre) => (
          <GenreTag
            key={genre}
            genre={genre}
            onRemove={isEditable ? () => onRemove(genre) : undefined}
          />
        ))}
      </div>
      {isEditable ? (
        <GenreSearchList genres={genres} onAdd={onAdd} onRemove={onRemove} />
      ) : null}
    </div>
  );
}

export function GenreRow({
  genres,
  isEditable,
  isChanged,
  onAdd,
  onRemove,
}: GenreEditActions & {
  genres: string[];
  isEditable: boolean;
  isChanged: boolean;
}) {
  const rowRef = useRef<HTMLDivElement>(null);
  const addTriggerRef = useRef<HTMLButtonElement>(null);
  const moreTriggerRef = useRef<HTMLButtonElement>(null);
  const [openPopover, setOpenPopover] = useState<GenrePopover | null>(null);
  const visibleGenres = genres.slice(0, VISIBLE_GENRE_COUNT);
  const moreCount = Math.max(genres.length - VISIBLE_GENRE_COUNT, 0);

  const closePopover = (shouldReturnFocus: boolean) => {
    const trigger = openPopover === 'add' ? addTriggerRef : moreTriggerRef;
    setOpenPopover(null);
    if (shouldReturnFocus) {
      trigger.current?.focus();
    }
  };
  usePopoverDismiss(rowRef, openPopover !== null, closePopover);

  const togglePopover = (popover: GenrePopover) =>
    setOpenPopover((current) => (current === popover ? null : popover));

  return (
    <div
      ref={rowRef}
      className={`relative mx-4.5 mb-3.5 flex h-10 flex-nowrap items-center gap-1.5 rounded-[10px] pr-1.5 pl-2.5 ${
        isChanged
          ? 'bg-[#fffaf0] shadow-[inset_0_0_0_1.5px_#c9973f]'
          : 'bg-[#f7f4f0]'
      }`}
    >
      {genres.length === 0 ? (
        <span className="truncate text-xs text-[#6b645d] italic">
          No genres yet
        </span>
      ) : (
        visibleGenres.map((genre) => (
          <GenreTag
            key={genre}
            genre={genre}
            className="max-w-30"
            onRemove={isEditable ? () => onRemove(genre) : undefined}
          />
        ))
      )}
      {moreCount > 0 ? (
        <button
          ref={moreTriggerRef}
          type="button"
          aria-expanded={openPopover === 'all'}
          aria-label={`Show all ${genres.length} genres`}
          onClick={() => togglePopover('all')}
          className="h-7 shrink-0 cursor-pointer rounded-full border-[1.5px] border-[#3d3935] bg-white px-2.5 text-[11px] font-extrabold text-[#3d3935]"
        >
          +{moreCount}
        </button>
      ) : null}
      {isEditable ? (
        <button
          ref={addTriggerRef}
          type="button"
          aria-expanded={openPopover === 'add'}
          aria-label="Add genre"
          onClick={() => togglePopover('add')}
          className={`${addGenreButtonClassName} ml-auto`}
        >
          + Genre
        </button>
      ) : null}
      {openPopover === 'add' ? (
        <AddGenrePopover
          genres={genres}
          className="top-11 right-1.5"
          onAdd={onAdd}
          onRemove={onRemove}
        />
      ) : null}
      {openPopover === 'all' ? (
        <AllGenresPopover
          genres={genres}
          isEditable={isEditable}
          onAdd={onAdd}
          onRemove={onRemove}
          onClose={() => closePopover(true)}
        />
      ) : null}
    </div>
  );
}

export function GenreChipList({
  genres,
  onAdd,
  onRemove,
}: GenreEditActions & { genres: string[] }) {
  const containerRef = useRef<HTMLDivElement>(null);
  const addTriggerRef = useRef<HTMLButtonElement>(null);
  const [isAddOpen, setIsAddOpen] = useState(false);

  const closeAddPopover = (shouldReturnFocus: boolean) => {
    setIsAddOpen(false);
    if (shouldReturnFocus) {
      addTriggerRef.current?.focus();
    }
  };
  usePopoverDismiss(containerRef, isAddOpen, closeAddPopover);

  return (
    <div ref={containerRef} className="flex flex-wrap items-center gap-1.5">
      {genres.map((genre) => (
        <GenreTag key={genre} genre={genre} onRemove={() => onRemove(genre)} />
      ))}
      <span className="relative">
        <button
          ref={addTriggerRef}
          type="button"
          aria-expanded={isAddOpen}
          aria-label="Add genre"
          onClick={() => setIsAddOpen((isOpen) => !isOpen)}
          className={addGenreButtonClassName}
        >
          + Genre
        </button>
        {isAddOpen ? (
          <AddGenrePopover
            genres={genres}
            className="top-full left-0 mt-1"
            onAdd={onAdd}
            onRemove={onRemove}
          />
        ) : null}
      </span>
    </div>
  );
}
