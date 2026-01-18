// components
import Button from '@/components/ui/Button';

// context
import { useDatabasePageContext } from 'lib/context/DatabasePageContext';

export default function PageSelector() {
  const {
    databaseItems: { min, max, total },
    page,
    setPage,
    handleGetMedia,
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
          label={'<- Prev Page'}
          width={125}
          fontSize={24}
          disabled={page === 1 || total === 0}
          onClick={() => {
            setPage((prev: number) => prev - 1);
            handleGetMedia();
          }}
        />
        <Button
          variant="primary"
          disabled={page === numPages || total === 0}
          width={125}
          fontSize={24}
          label={'Next Page ->'}
          onClick={() => {
            setPage((prev: number) => prev + 1);
            handleGetMedia();
          }}
        />
      </div>
    </div>
  );
}
