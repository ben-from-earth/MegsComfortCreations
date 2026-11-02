'use client';

import PageSelector from '@/components/shared/page-selector';
import LibraryItem from '@/library/library-item';
import { useLibraryPageContext } from 'lib/context/library-page-context';

export default function LibraryItems() {
  const { libraryList } = useLibraryPageContext();

  return (
    <div className='border-darkpink bg-lightpink relative mt-2.5 flex min-w-lg flex-col items-center gap-2 self-center rounded-lg border-3 p-2 font-["Just_Another_Hand"] text-2xl tracking-wider shadow-[5px_5px_30px_rgba(0,0,0,0.3)]'>
      <PageSelector />
      {libraryList.items.map((item) => {
        return (
          <LibraryItem key={`${libraryList.type}-${item.id}`} info={item} />
        );
      })}
    </div>
  );
}
