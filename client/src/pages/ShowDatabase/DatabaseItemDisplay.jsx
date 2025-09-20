import PageChooser from '@/components/PageChooser';
import DatabaseItem from './DatabaseItem';
import { useContext } from 'react';
import DatabasePageContext from '@/context/DatabasePageContext';

const DatabaseItemDisplay = () => {
  const { databaseItems } = useContext(DatabasePageContext);
  return (
    <div className='border-3 w-4/10 relative mt-2.5 flex flex-col items-center gap-2 rounded-lg border-[var(--darkpink)] bg-[var(--lightpink)] p-2 font-["Just_Another_Hand"] text-2xl tracking-wider shadow-[5px_5px_30px_rgba(0,0,0,0.3)]'>
      <PageChooser />
      {databaseItems.items.map((item) => {
        return (
          <DatabaseItem key={item.id} info={item} type={databaseItems.type} />
        );
      })}
    </div>
  );
};

export default DatabaseItemDisplay;
