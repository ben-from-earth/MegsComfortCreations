// components
import PageSelector from '@/shared/PageSelector';
import DatabaseItem from '@//showdatabase/DatabaseItem';

// context
import { useDatabasePageContext } from 'lib/context/DatabasePageContext';

export default function DatabaseItemsContainer() {
  const { databaseItems } = useDatabasePageContext();
  return (
    <div className='border-darkpink bg-lightpink relative mt-2.5 flex min-w-lg flex-col items-center gap-2 rounded-lg border-3 p-2 font-["Just_Another_Hand"] text-2xl tracking-wider shadow-[5px_5px_30px_rgba(0,0,0,0.3)]'>
      <PageSelector />
      {databaseItems.items.map((item) => {
        return <DatabaseItem key={item.id} info={item} />;
      })}
    </div>
  );
}
