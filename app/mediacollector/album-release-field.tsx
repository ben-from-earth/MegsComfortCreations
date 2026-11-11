import { useEffect, useId, useRef, useState } from 'react';
import { isoToUsDate, isValidUsDate, usDateToIso } from 'lib/helpers/us-date';
import { editStateClassNames } from '@/mediacollector/collected-item-text-field';

const WEEKDAY_LABELS = ['S', 'M', 'T', 'W', 'T', 'F', 'S'] as const;

interface ReleaseMonth {
  year: number;
  month: number;
}

function monthFromValue(value: string | null): ReleaseMonth {
  const usDate = isoToUsDate(value);
  if (!value || usDate === '') {
    const today = new Date();
    return { year: today.getFullYear(), month: today.getMonth() + 1 };
  }
  return {
    year: Number(value.slice(0, 4)),
    month: Number(value.slice(5, 7)),
  };
}

function shiftMonth(current: ReleaseMonth, offset: number): ReleaseMonth {
  const date = new Date(current.year, current.month - 1 + offset, 1);
  return { year: date.getFullYear(), month: date.getMonth() + 1 };
}

function isoFromParts(year: number, month: number, day: number) {
  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function monthName(current: ReleaseMonth) {
  return new Intl.DateTimeFormat('en-US', { month: 'long' }).format(
    new Date(current.year, current.month - 1, 1),
  );
}

function dayLabel(year: number, month: number, day: number) {
  return new Intl.DateTimeFormat('en-US', {
    month: 'long',
    day: 'numeric',
    year: 'numeric',
  }).format(new Date(year, month - 1, day));
}

const stepButtonClassName =
  'flex h-7 w-7 cursor-pointer items-center justify-center rounded-lg border-0 bg-transparent p-0 text-sm font-extrabold text-[#4a443f]';

export interface AlbumReleaseFieldProps {
  value: string | null;
  onChange: (value: string | null) => void;
  readOnly?: boolean;
  changed?: boolean;
  isOpen: boolean;
  onOpenChange: (isOpen: boolean) => void;
}

export default function AlbumReleaseField({
  value,
  onChange,
  readOnly = false,
  changed = false,
  isOpen,
  onOpenChange,
}: AlbumReleaseFieldProps) {
  const inputId = useId();
  const dialogId = useId();
  const rootRef = useRef<HTMLDivElement>(null);
  const buttonRef = useRef<HTMLButtonElement>(null);
  const [text, setText] = useState(() => isoToUsDate(value));
  const [view, setView] = useState(() => monthFromValue(value));
  const invalid = text.trim() !== '' && !isValidUsDate(text);
  const shownText = readOnly ? isoToUsDate(value) : text;

  useEffect(() => {
    setText(isoToUsDate(value));
  }, [value, readOnly]);

  useEffect(() => {
    if (readOnly && isOpen) {
      onOpenChange(false);
    }
  }, [readOnly, isOpen, onOpenChange]);

  useEffect(() => {
    if (!isOpen) {
      return;
    }
    const root = rootRef.current;
    const handlePointerDown = (event: PointerEvent) => {
      if (event.target instanceof Node && root?.contains(event.target)) {
        return;
      }
      onOpenChange(false);
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') {
        return;
      }
      event.preventDefault();
      onOpenChange(false);
      buttonRef.current?.focus();
    };
    document.addEventListener('pointerdown', handlePointerDown);
    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('pointerdown', handlePointerDown);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [isOpen, onOpenChange]);

  const togglePicker = () => {
    if (readOnly) {
      return;
    }
    if (isOpen) {
      onOpenChange(false);
      return;
    }
    setView(monthFromValue(value));
    onOpenChange(true);
  };

  const handleTextChange = (next: string) => {
    if (readOnly) {
      return;
    }
    if (next.trim() === '') {
      setText('');
      onChange(null);
      return;
    }
    setText(next);
    const iso = usDateToIso(next);
    if (iso) {
      onChange(iso);
    }
  };

  const selectDay = (day: number) => {
    const iso = isoFromParts(view.year, view.month, day);
    onChange(iso);
    setText(isoToUsDate(iso));
  };

  const leadingBlankCount = new Date(view.year, view.month - 1, 1).getDay();
  const dayCount = new Date(view.year, view.month, 0).getDate();
  const storedUsDate = isoToUsDate(value);

  return (
    <div
      ref={rootRef}
      className="relative flex min-w-0 flex-col gap-1 text-xs font-bold text-[#6b645d]"
    >
      <label htmlFor={inputId}>Released</label>
      <span className="relative flex">
        <input
          id={inputId}
          type="text"
          inputMode="numeric"
          placeholder="MM/DD/YYYY"
          readOnly={readOnly}
          aria-invalid={invalid}
          autoComplete="off"
          className={`h-9.5 w-full min-w-0 rounded-lg border pr-10 pl-2.5 text-sm font-normal text-[#2b2826] ${
            readOnly
              ? editStateClassNames.readOnly
              : invalid
                ? 'border-[#b0533a] bg-[#fcfbf9] ring-1 ring-[#b0533a]'
                : editStateClassNames[changed ? 'changed' : 'editable']
          }`}
          value={shownText}
          onChange={(event) => handleTextChange(event.target.value)}
        />
        <button
          ref={buttonRef}
          type="button"
          disabled={readOnly}
          aria-label="Choose release date"
          aria-haspopup="dialog"
          aria-expanded={isOpen}
          aria-controls={dialogId}
          onClick={togglePicker}
          className={`absolute top-1 right-1 flex h-7.5 w-7.5 items-center justify-center rounded-md border-0 bg-transparent p-0 ${
            readOnly
              ? 'cursor-default text-[#b3aaa0]'
              : 'cursor-pointer text-[#4a443f]'
          }`}
        >
          <svg
            width="16"
            height="16"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
            strokeLinecap="round"
            strokeLinejoin="round"
            aria-hidden="true"
          >
            <rect x="3.5" y="5" width="17" height="15" rx="2" />
            <path d="M3.5 10h17M8 3v4M16 3v4" />
          </svg>
        </button>
        {isOpen ? (
          <div
            id={dialogId}
            role="dialog"
            aria-label="Choose release date"
            className="absolute top-[calc(100%+6px)] right-0 z-30 box-border flex w-[272px] flex-col gap-2 rounded-xl border border-[#e0d9d0] bg-white p-3 shadow-[0_12px_28px_rgba(40,35,30,0.18)]"
          >
            <div className="flex items-center justify-between">
              <button
                type="button"
                aria-label="Previous year"
                onClick={() =>
                  setView((current) => ({
                    year: current.year - 1,
                    month: current.month,
                  }))
                }
                className={stepButtonClassName}
              >
                «
              </button>
              <span className="text-sm font-extrabold text-[#2b2826]">
                {view.year}
              </span>
              <button
                type="button"
                aria-label="Next year"
                onClick={() =>
                  setView((current) => ({
                    year: current.year + 1,
                    month: current.month,
                  }))
                }
                className={stepButtonClassName}
              >
                »
              </button>
            </div>
            <div className="flex items-center justify-between">
              <button
                type="button"
                aria-label="Previous month"
                onClick={() => setView((current) => shiftMonth(current, -1))}
                className={stepButtonClassName}
              >
                ‹
              </button>
              <span className="text-sm font-extrabold text-[#2b2826]">
                {monthName(view)}
              </span>
              <button
                type="button"
                aria-label="Next month"
                onClick={() => setView((current) => shiftMonth(current, 1))}
                className={stepButtonClassName}
              >
                ›
              </button>
            </div>
            <div
              aria-hidden="true"
              className="grid grid-cols-7 text-center text-[11px] font-extrabold text-[#6b645d]"
            >
              {WEEKDAY_LABELS.map((label, index) => (
                <span key={`${label}-${index}`}>{label}</span>
              ))}
            </div>
            <div className="grid grid-cols-7">
              {Array.from({ length: leadingBlankCount }, (_, index) => (
                <span key={`blank-${index}`} className="h-8" />
              ))}
              {Array.from({ length: dayCount }, (_, index) => {
                const day = index + 1;
                const iso = isoFromParts(view.year, view.month, day);
                const isSelected = value === iso;
                return (
                  <button
                    key={iso}
                    type="button"
                    aria-label={dayLabel(view.year, view.month, day)}
                    aria-pressed={isSelected}
                    onClick={() => selectDay(day)}
                    className={`h-8 cursor-pointer rounded-lg border-0 text-sm font-bold ${
                      isSelected
                        ? 'bg-[#2b2826] text-white'
                        : 'bg-transparent text-[#2b2826] hover:bg-[#f7f4f0]'
                    }`}
                  >
                    {day}
                  </button>
                );
              })}
            </div>
            <div className="flex items-center justify-between gap-2 border-t border-[#efe9e2] pt-2">
              <span className="text-sm font-bold text-[#2b2826]">
                {storedUsDate}
              </span>
              <button
                type="button"
                onClick={() => onOpenChange(false)}
                className="h-8 cursor-pointer rounded-lg border border-[#e0d9d0] bg-white px-3 text-[13px] font-bold text-[#2b2826]"
              >
                Done
              </button>
            </div>
          </div>
        ) : null}
      </span>
    </div>
  );
}
