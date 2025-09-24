import Button from '@/components/Button';
import { titleRearrange } from '@/pages/MediaCollector/helpers/mediaCollectorHelpers';

const DatabaseSavedWidget = ({ data, close }) => {
  return (
    <div className='z-100 border-3 max-h-100 fixed left-1/2 top-1/2 flex w-4/12 -translate-x-1/2 -translate-y-1/2 flex-col content-center items-center overflow-y-auto rounded-md border-[var(--darkpink)] bg-[var(--lightpink)] p-2 font-["Just_Another_Hand"] text-2xl tracking-wider text-black'>
      Results of saving to database:
      <ol className="list-decimal">
        {data.map((item, idx) => (
          <li key={idx}>
            {!item.actionCompleted ? (
              <>
                {`Error saving ${
                  item.saveAttemptItem.title
                    ? titleRearrange(item.saveAttemptItem.title)
                    : '[Missing Title]'
                }:`}
                <ol className="pl-7">
                  {item.errors.map((error, eIdx) => (
                    <li type="a" key={eIdx}>
                      {error}
                    </li>
                  ))}
                </ol>
              </>
            ) : (
              item.message
            )}
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
};

export default DatabaseSavedWidget;
