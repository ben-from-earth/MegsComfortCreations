// react, redux imports
import { Dispatch, SetStateAction } from 'react';

// components
import Button from '@/components/ui/button';

// helpers
import { titleRearrange } from 'lib/helpers/title-rearrange';

export interface AreYouSureProps {
  setAreYouSure: Dispatch<SetStateAction<boolean>>;
  onDelete: () => Promise<void>;
  title: string;
}

export default function AreYouSure({
  setAreYouSure,
  onDelete,
  title,
}: AreYouSureProps) {
  return (
    <div className="border-darkpink bg-lightpink fixed top-1/2 left-1/2 z-100 flex h-1/4 w-1/4 -translate-x-1/2 -translate-y-1/2 flex-col content-center items-center justify-center gap-10 rounded-md border-3 p-2 text-2xl tracking-wider text-black">
      <p className="text-center">
        Are you sure you want to delete
        {title ? titleRearrange(title) : '[Missing Title]'}?
      </p>
      <div className="flex gap-2">
        <Button
          variant="primary"
          label={'Yes'}
          width={75}
          fontSize={30}
          onClick={() => onDelete()}
          className="bg-emerald-300 hover:bg-green-400"
        />
        <Button
          variant="primary"
          label={'No'}
          width={75}
          fontSize={30}
          className="bg-red-300 hover:bg-red-400"
          onClick={() => setAreYouSure(false)}
        />
      </div>
    </div>
  );
}
