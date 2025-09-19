import DatabaseItem from './DatabaseItem';

const DatabaseItemDisplay = ({ databaseItems, handleGetMedia }) => {
  return (
    <div className='border-3 w-4/10 mt-2.5 flex flex-col items-center gap-2 rounded-lg border-[var(--darkpink)] bg-[var(--lightpink)] p-2 font-["Just_Another_Hand"] text-2xl tracking-wider shadow-[5px_5px_30px_rgba(0,0,0,0.3)]'>
      <p>
        Showing {databaseItems.min}-
        {Math.min(databaseItems.max, databaseItems.total)} of{' '}
        {databaseItems.total}
      </p>
      {databaseItems.items.map((item) => {
        return (
          <DatabaseItem
            key={item.id}
            info={item}
            type={databaseItems.type}
            handleGetMedia={handleGetMedia}
          />
        );
      })}
    </div>
  );
};

export default DatabaseItemDisplay;
