// components
import Button from '@/shared/Button';

// interfaces and types
import { DatabaseSaveServerResponse } from 'lib/interfaces/globalInterfaces';

export interface DatabaseSavedWidgetProps {
  data: DatabaseSaveServerResponse;
  close: () => void;
}

export default function DatabaseSavedWidget({
  data,
  close,
}: DatabaseSavedWidgetProps) {
  const totalCount = data.length;
  const failedItems = data.filter((item) => 'error' in item);
  const successCount = totalCount - failedItems.length;
  return (
    <div className="border-darkpink bg-lightpink fixed top-1/2 left-1/2 z-100 flex max-h-100 w-5/12 -translate-x-1/2 -translate-y-1/2 flex-col content-center items-center overflow-y-auto rounded-md border-3 p-2 text-2xl tracking-wider text-black">
      Results of saving to database:
      <p>Successful saves: {successCount}</p>
      <p>Failed saves: {failedItems.length}</p>
      <ol className="list-decimal">
        {failedItems.map((item, idx) => (
          <li key={idx}>
            {`Error saving ${item.title} to database`}
            <ol type="a" className="pl-7">
              {item.errors.map((error, eIdx) => (
                <li key={eIdx}>{error}</li>
              ))}
            </ol>
          </li>
        ))}
      </ol>
      <Button
        additionalStyling={'absolute right-2 top-2'}
        onClick={close}
        label={'Close'}
        width={100}
        fontSize={25}
      />
    </div>
  );
}
