// components
import ArrowBackIcon from '@mui/icons-material/ArrowBack';
import ArrowForwardIcon from '@mui/icons-material/ArrowForward';
import Button from '@/components/ui/Button';

// context
import { useDatabasePageContext } from 'lib/context/DatabasePageContext';

export default function PageSelector() {
  const {
    databaseItems: { min, max, total },
    page,
    setPage,
  } = useDatabasePageContext();
  const numPages = Math.ceil(total / (max - min + 1));
  return (
    <div className="flex w-full">
      <p className="mr-auto pt-0.5">
        Showing {Math.min(min, total)}-{Math.min(max, total)} of {total}
      </p>
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
