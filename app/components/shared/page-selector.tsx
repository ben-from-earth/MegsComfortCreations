import ArrowBackIcon from '@mui/icons-material/ArrowBack';
import ArrowForwardIcon from '@mui/icons-material/ArrowForward';
import Button from '@/components/ui/button';
import { useLibraryPageContext } from 'lib/context/library-page-context';

export default function PageSelector() {
  const {
    libraryList: { min, max, total },
    page,
    setPage,
  } = useLibraryPageContext();
  const numPages = Math.ceil(total / (max - min + 1));
  return (
    <div className="flex w-full">
      <div className="mr-auto flex items-center gap-2">
        <p className="pt-0.5">
          Showing {Math.min(min, total)}-{Math.min(max, total)} of {total}
        </p>
      </div>
      <div className="ml-auto flex gap-2">
        <Button
          variant="primary"
          width={44}
          className="flex items-center justify-center"
          disabled={page === 1 || total === 0}
          onClick={() => {
            setPage((prev: number) => prev - 1);
          }}
        >
          <ArrowBackIcon />
        </Button>
        <Button
          variant="primary"
          width={44}
          className="flex items-center justify-center"
          disabled={page === numPages || total === 0}
          onClick={() => {
            setPage((prev: number) => prev + 1);
          }}
        >
          <ArrowForwardIcon />
        </Button>
      </div>
    </div>
  );
}
