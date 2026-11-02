import { useLibraryPageContext } from 'lib/context/library-page-context';
import { getPageWindow } from './page-window';

const pagerButtonClassName =
  'h-9.5 min-w-9.5 cursor-pointer rounded-full border-0 px-3 font-bold disabled:cursor-not-allowed disabled:text-[#b8b0a6]';

export default function LibraryPager() {
  const {
    page,
    setPage,
    limit,
    libraryList: { total },
  } = useLibraryPageContext();
  const pageCount = Math.ceil(total / limit);

  return (
    <nav
      aria-label="Pages"
      className="flex flex-wrap gap-1 rounded-full border border-[#e6e0d8] bg-white p-1"
    >
      <button
        type="button"
        className={`${pagerButtonClassName} bg-transparent text-[#2b2826]`}
        disabled={page <= 1}
        onClick={() => setPage(page - 1)}
      >
        Previous
      </button>
      {getPageWindow(page, pageCount).map((pageNumber, index) =>
        pageNumber === 'ellipsis' ? (
          <span
            key={`ellipsis-${index}`}
            className="flex h-9.5 min-w-6 items-center justify-center font-bold text-[#6b645d]"
          >
            …
          </span>
        ) : (
          <button
            key={pageNumber}
            type="button"
            aria-current={pageNumber === page ? 'page' : undefined}
            className={`${pagerButtonClassName} ${
              pageNumber === page
                ? 'bg-[#2a4637] text-white'
                : 'bg-transparent text-[#2b2826]'
            }`}
            onClick={() => setPage(pageNumber)}
          >
            {pageNumber}
          </button>
        ),
      )}
      <button
        type="button"
        className={`${pagerButtonClassName} bg-transparent text-[#2b2826]`}
        disabled={page >= pageCount}
        onClick={() => setPage(page + 1)}
      >
        Next
      </button>
    </nav>
  );
}
