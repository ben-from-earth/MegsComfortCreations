import { useEffect, useId, useRef, useState } from 'react';
import { isoToUsDate, isValidUsDate, usDateToIso } from 'lib/helpers/us-date';
import { editStateClassNames } from '@/mediacollector/collected-item-text-field';

const WEEKDAY_LABELS = ['S', 'M', 'T', 'W', 'T', 'F', 'S'] as const;

export const releasePickerOpenCardClassName =
  'has-[[data-release-picker]]:z-20 has-[[data-release-picker]]:overflow-visible';

interface ReleaseMonth {
  year: number;
  month: number;
}

interface ReleaseDraft {
  text: string;
  typedOver: string | null;
}

function monthFromValue(value: string | null): ReleaseMonth {
  if (!value || isoToUsDate(value) === '') {
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

const monthNameFormat = new Intl.DateTimeFormat('en-US', { month: 'long' });
const dayLabelFormat = new Intl.DateTimeFormat('en-US', {
  month: 'long',
  day: 'numeric',
  year: 'numeric',
});

const stepButtonClassName =
  'flex h-7 w-7 shrink-0 cursor-pointer items-center justify-center rounded-lg border-0 bg-transparent p-0 text-base leading-none font-extrabold text-[#4a443f] hover:bg-[#f7f4f0]';

function CalendarIcon() {
  return (
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
  );
}

export interface AlbumReleaseFieldProps {
  value: string | null;
  onChange: (value: string | null) => void;
  readOnly: boolean;
  changed: boolean;
}

export default function AlbumReleaseField({
  value,
  onChange,
  readOnly,
  changed,
}: AlbumReleaseFieldProps) {
  const inputId = useId();
  const dialogId = useId();
  const rootRef = useRef<HTMLDivElement>(null);
  const buttonRef = useRef<HTMLButtonElement>(null);
  const [isOpen, setIsOpen] = useState(false);
  const [draft, setDraft] = useState<ReleaseDraft | null>(null);
  const [view, setView] = useState(() => monthFromValue(value));
  const storedText = isoToUsDate(value);
  const text =
    !readOnly && draft !== null && draft.typedOver === value
      ? draft.text
      : storedText;
  const isInvalid = text.trim() !== '' && !isValidUsDate(text);
  const isPickerOpen = isOpen && !readOnly;

  useEffect(() => {
    if (!isPickerOpen) {
      return;
    }
    const handlePointerDown = (event: PointerEvent) => {
      if (
        event.target instanceof Node &&
        rootRef.current?.contains(event.target)
      ) {
        return;
      }
      setIsOpen(false);
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') {
        return;
      }
      event.preventDefault();
      setIsOpen(false);
      buttonRef.current?.focus();
    };
    document.addEventListener('pointerdown', handlePointerDown);
    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('pointerdown', handlePointerDown);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [isPickerOpen]);

  const togglePicker = () => {
    if (!isPickerOpen) {
      setView(monthFromValue(value));
    }
    setIsOpen(!isPickerOpen);
  };

  const handleTextChange = (next: string) => {
    if (next.trim() === '') {
      setDraft(null);
      onChange(null);
      return;
    }
    const iso = usDateToIso(next);
    if (iso) {
      setDraft(null);
      onChange(iso);
      return;
    }
    setDraft({ text: next, typedOver: value });
  };

  const selectDay = (day: number) => {
    setDraft(null);
    onChange(isoFromParts(view.year, view.month, day));
  };

  const leadingBlankCount = new Date(view.year, view.month - 1, 1).getDay();
  const dayCount = new Date(view.year, view.month, 0).getDate();
  const viewDate = new Date(view.year, view.month - 1, 1);

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
          autoComplete="off"
          readOnly={readOnly}
          aria-invalid={isInvalid}
          className={`h-9.5 w-full min-w-0 rounded-lg border pr-10 pl-2.5 text-sm font-normal text-[#2b2826] ${
            readOnly
              ? editStateClassNames.readOnly
              : isInvalid
                ? 'border-[#b0533a] bg-[#fcfbf9] ring-1 ring-[#b0533a]'
                : editStateClassNames[changed ? 'changed' : 'editable']
          }`}
          value={text}
          onChange={(event) => handleTextChange(event.target.value)}
        />
        <button
          ref={buttonRef}
          type="button"
          disabled={readOnly}
          aria-label="Choose release date"
          aria-haspopup="dialog"
          aria-expanded={isPickerOpen}
          aria-controls={isPickerOpen ? dialogId : undefined}
          onClick={togglePicker}
          className={`absolute top-1 right-1 flex h-7.5 w-7.5 items-center justify-center rounded-md border-0 bg-transparent p-0 ${
            readOnly
              ? 'cursor-default text-[#b3aaa0]'
              : 'cursor-pointer text-[#4a443f] hover:bg-[#f7f4f0]'
          }`}
        >
          <CalendarIcon />
        </button>
        {isPickerOpen ? (
          <div
            id={dialogId}
            role="dialog"
            aria-label="Choose release date"
            data-release-picker
            className="absolute top-[calc(100%+6px)] right-0 z-30 box-border flex w-68 flex-col gap-2 rounded-xl border border-[#e6e0d8] bg-white p-3 shadow-[0_12px_28px_rgba(40,35,30,0.18)]"
          >
            <div className="flex items-center gap-0.5">
              <button
                type="button"
                aria-label="Previous year"
                onClick={() =>
                  setView((current) => ({ ...current, year: current.year - 1 }))
                }
                className={stepButtonClassName}
              >
                «
              </button>
              <button
                type="button"
                aria-label="Previous month"
                onClick={() => setView((current) => shiftMonth(current, -1))}
                className={stepButtonClassName}
              >
                ‹
              </button>
              <span className="flex-1 text-center text-sm font-extrabold text-[#2b2826]">
                {monthNameFormat.format(viewDate)} {view.year}
              </span>
              <button
                type="button"
                aria-label="Next month"
                onClick={() => setView((current) => shiftMonth(current, 1))}
                className={stepButtonClassName}
              >
                ›
              </button>
              <button
                type="button"
                aria-label="Next year"
                onClick={() =>
                  setView((current) => ({ ...current, year: current.year + 1 }))
                }
                className={stepButtonClassName}
              >
                »
              </button>
            </div>
            <div
              aria-hidden="true"
              className="grid grid-cols-7 text-center text-[11px] font-extrabold text-[#6b645d]"
            >
              {WEEKDAY_LABELS.map((label, index) => (
                <span key={index}>{label}</span>
              ))}
            </div>
            <div className="grid grid-cols-7 gap-y-0.5">
              {Array.from({ length: leadingBlankCount }, (_, index) => (
                <span key={`blank-${index}`} />
              ))}
              {Array.from({ length: dayCount }, (_, index) => {
                const day = index + 1;
                const iso = isoFromParts(view.year, view.month, day);
                const isSelected = value === iso;
                return (
                  <button
                    key={iso}
                    type="button"
                    aria-label={dayLabelFormat.format(
                      new Date(view.year, view.month - 1, day),
                    )}
                    aria-pressed={isSelected}
                    onClick={() => selectDay(day)}
                    className={`h-8 cursor-pointer rounded-lg border-0 text-[13px] font-bold ${
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
              <span className="text-[13px] font-bold text-[#2b2826]">
                {storedText || 'No date'}
              </span>
              <button
                type="button"
                onClick={() => setIsOpen(false)}
                className="h-8 cursor-pointer rounded-lg border-0 bg-[#2b2826] px-3.5 text-[13px] font-extrabold text-white"
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
