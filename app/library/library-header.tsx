import { use, useState } from 'react';
import AddIcon from '@mui/icons-material/Add';
import SearchIcon from '@mui/icons-material/Search';
import GenreContext from 'lib/context/genre-context';
import { useLibraryPageContext } from 'lib/context/library-page-context';
import { MEDIA_TYPE_DEFINITIONS, MediaType } from 'lib/constants/media-types';
import { NO_GENRE_FILTER } from '@/lib/enums/genre-enums';
import {
  BOOK_SORT_SELECT_OPTIONS,
  NON_BOOK_SORT_SELECT_OPTIONS,
} from 'lib/constants/database-sort-options';
import AddBookDialog from '@/library/add-book-dialog';

const LIBRARY_PAGE_SIZES = [12, 24, 48] as const;

const labelClassName =
  'flex flex-col gap-1.5 text-[13px] font-bold text-[#6b645d]';
const inputClassName =
  'h-11 rounded-[10px] border border-[#ddd5cb] bg-[#fcfbf9] px-3 text-[15px] text-[#2b2826] disabled:cursor-not-allowed disabled:opacity-50';
const primaryButtonClassName =
  'flex h-11 cursor-pointer items-center justify-center gap-2 rounded-xl border border-transparent bg-[#2a4637] px-5 text-base font-extrabold text-white';

export default function LibraryHeader() {
  const {
    type,
    setType,
    limit,
    setLimit,
    sortBy,
    setSortBy,
    setPage,
    genre,
    setGenre,
    ascDesc,
    setAscDesc,
    setTitleSearch,
    setYear,
    missingYear,
    setMissingYear,
    setArtist,
    releaseDateFrom,
    setReleaseDateFrom,
    releaseDateTo,
    setReleaseDateTo,
  } = useLibraryPageContext();
  const genres = use(GenreContext);
  const [titleInput, setTitleInput] = useState('');
  const [yearInput, setYearInput] = useState('');
  const [artistInput, setArtistInput] = useState('');
  const [isAddBookOpen, setIsAddBookOpen] = useState(false);
  const isBook = type === 'book';
  const hasYearFilter = type === 'movie' || type === 'videoGame';
  const sortOptions = isBook
    ? BOOK_SORT_SELECT_OPTIONS
    : NON_BOOK_SORT_SELECT_OPTIONS;

  const handleTypeChange = (mediaType: MediaType) => {
    setPage(1);
    setType(mediaType);
    if (mediaType !== 'book') {
      setSortBy('title');
      setGenre('');
    }
    setYearInput('');
    setYear(undefined);
    setMissingYear(false);
    setArtistInput('');
    setArtist('');
    setReleaseDateFrom('');
    setReleaseDateTo('');
  };

  return (
    <section
      aria-label="Library filters"
      className="overflow-hidden rounded-[20px] border border-[#e6e0d8] bg-white shadow-[0_8px_30px_rgba(60,50,40,0.07)]"
    >
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-[#e3e9de] bg-[#f3f6f0] px-7 py-4.5">
        <h1 className="m-0 font-(family-name:--font-hand) text-[52px] leading-[0.9] font-normal tracking-[1px]">
          Meg’s Library
        </h1>
        {isBook ? (
          <button
            type="button"
            className={primaryButtonClassName}
            onClick={() => setIsAddBookOpen(true)}
          >
            <AddIcon fontSize="small" />
            Add book
          </button>
        ) : null}
      </div>

      <div className="flex flex-col gap-5 px-7 pt-6 pb-7">
        <div className="flex flex-wrap gap-2">
          {MEDIA_TYPE_DEFINITIONS.map(({ mediaType, plural }) => {
            const isActive = type === mediaType;
            return (
              <button
                key={mediaType}
                type="button"
                aria-pressed={isActive}
                onClick={() => handleTypeChange(mediaType)}
                className={`h-9 cursor-pointer rounded-full border-2 px-3.5 text-[13px] font-extrabold ${
                  isActive
                    ? 'border-[#2a4637] bg-[#2a4637] text-white'
                    : 'border-[#e0d9d0] bg-white text-[#6b645d]'
                }`}
              >
                {plural}
              </button>
            );
          })}
        </div>

        <form
          className="flex flex-wrap items-end gap-4"
          onSubmit={(event) => {
            event.preventDefault();
            setPage(1);
            setTitleSearch(titleInput);
            setYear(yearInput === '' ? undefined : Number(yearInput));
            setArtist(artistInput.trim());
          }}
        >
          <label className={`${labelClassName} flex-[1_1_260px]`}>
            Title
            <span className="relative flex items-center">
              <SearchIcon
                fontSize="small"
                className="pointer-events-none absolute left-3 text-[#6b645d]"
              />
              <input
                type="search"
                value={titleInput}
                onChange={(event) => setTitleInput(event.target.value)}
                placeholder="Search titles"
                className={`${inputClassName} w-full pl-9`}
              />
            </span>
          </label>

          {isBook ? (
            <label className={labelClassName}>
              Genre
              <select
                value={genre}
                onChange={(event) => {
                  const { value } = event.target;
                  setPage(1);
                  setGenre(
                    genres.find((option) => option === value) ??
                      (value === NO_GENRE_FILTER ? NO_GENRE_FILTER : ''),
                  );
                }}
                className={inputClassName}
              >
                <option value="">All genres</option>
                <option value={NO_GENRE_FILTER}>No genre</option>
                {genres.map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
            </label>
          ) : null}

          {hasYearFilter ? (
            <>
              <label className={labelClassName}>
                Year
                <input
                  type="number"
                  min={1000}
                  max={9999}
                  step={1}
                  value={yearInput}
                  disabled={missingYear}
                  onChange={(event) => setYearInput(event.target.value)}
                  placeholder="Any year"
                  className={`${inputClassName} w-32`}
                />
              </label>
              <label className="flex h-11 cursor-pointer items-center gap-2 text-[13px] font-bold text-[#6b645d]">
                <input
                  type="checkbox"
                  checked={missingYear}
                  onChange={(event) => {
                    const { checked } = event.target;
                    setPage(1);
                    setMissingYear(checked);
                    if (checked) {
                      setYearInput('');
                      setYear(undefined);
                    }
                  }}
                  className="h-4 w-4 cursor-pointer accent-[#2a4637]"
                />
                Show items with no year
              </label>
            </>
          ) : null}

          {type === 'album' ? (
            <>
              <label className={`${labelClassName} flex-[1_1_200px]`}>
                Artist
                <input
                  type="search"
                  value={artistInput}
                  onChange={(event) => setArtistInput(event.target.value)}
                  placeholder="Search artists"
                  className={`${inputClassName} w-full`}
                />
              </label>
              <label className={labelClassName}>
                From
                <input
                  type="date"
                  value={releaseDateFrom}
                  onChange={(event) => {
                    setPage(1);
                    setReleaseDateFrom(event.target.value);
                  }}
                  className={inputClassName}
                />
              </label>
              <label className={labelClassName}>
                To
                <input
                  type="date"
                  value={releaseDateTo}
                  onChange={(event) => {
                    setPage(1);
                    setReleaseDateTo(event.target.value);
                  }}
                  className={inputClassName}
                />
              </label>
            </>
          ) : null}

          <label className={labelClassName}>
            Sort
            <span className="flex gap-2">
              <select
                value={sortBy}
                disabled={!isBook}
                onChange={(event) => {
                  const option = sortOptions.find(
                    ({ value }) => value === event.target.value,
                  );
                  if (!option) return;
                  setPage(1);
                  setSortBy(option.value);
                }}
                className={inputClassName}
              >
                {sortOptions.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                  </option>
                ))}
              </select>
              <button
                type="button"
                onClick={() => {
                  setPage(1);
                  setAscDesc((current) => (current === 'asc' ? 'desc' : 'asc'));
                }}
                className="h-11 min-w-16 cursor-pointer rounded-[10px] border border-[#ddd5cb] bg-white px-3 text-[15px] font-extrabold text-[#2b2826]"
              >
                {ascDesc === 'asc' ? 'A–Z' : 'Z–A'}
              </button>
            </span>
          </label>

          <label className={labelClassName}>
            Per page
            <select
              value={String(limit)}
              onChange={(event) => {
                const size = LIBRARY_PAGE_SIZES.find(
                  (option) => String(option) === event.target.value,
                );
                if (!size) return;
                setPage(1);
                setLimit(size);
              }}
              className={inputClassName}
            >
              {LIBRARY_PAGE_SIZES.map((size) => (
                <option key={size} value={size}>
                  {size}
                </option>
              ))}
            </select>
          </label>

          <button type="submit" className={primaryButtonClassName}>
            <SearchIcon fontSize="small" />
            Search
          </button>
        </form>
      </div>

      {isAddBookOpen ? (
        <AddBookDialog onClose={() => setIsAddBookOpen(false)} />
      ) : null}
    </section>
  );
}
