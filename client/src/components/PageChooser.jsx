import Button from '@/components/Button';
import { useContext } from 'react';
import DatabasePageContext from '@/context/DatabasePageContext';

const PageChooser = () => {
  const { databaseItems, page, setPage, handleGetMedia } =
    useContext(DatabasePageContext);
  const { min, max, total } = databaseItems;
  const numPages = Math.ceil(total / (max - min + 1));
  return (
    <div className="flex w-full">
      <p className="mr-auto pt-0.5">
        Showing {Math.min(min, total)}-{Math.min(max, total)} of {total}
      </p>
      <div className="ml-auto flex gap-2">
        <Button
          label={'<- Prev Page'}
          disabled={page === 1 || total === 0}
          onClick={() => {
            setPage((prev) => prev - 1);
            handleGetMedia();
          }}
        />
        <Button
          disabled={page === numPages || total === 0}
          label={'Next Page ->'}
          onClick={() => {
            setPage((prev) => prev + 1);
            handleGetMedia();
          }}
        />
      </div>
    </div>
  );
};

export default PageChooser;
