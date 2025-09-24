import Button from '@/components/Button';
import { titleRearrange } from '@/pages/MediaCollector/helpers/mediaCollectorHelpers';

// Are You Sure component for verification of deletion of an item from the database
const AreYouSure = ({ setAreYouSure, onDelete, title }) => {
  return (
    <div className='z-100 border-3 fixed left-1/2 top-1/2 flex h-1/4 w-1/4 -translate-x-1/2 -translate-y-1/2 flex-col content-center items-center justify-center gap-10 rounded-md border-[var(--darkpink)] bg-[var(--lightpink)] p-2 font-["Just_Another_Hand"] text-2xl tracking-wider text-black'>
      <p className="text-center">
        Are you sure you want to delete "
        {title ? titleRearrange(title) : '[Missing Title]'}"?
      </p>
      <div className="flex gap-2">
        <Button
          label={'Yes'}
          width={75}
          fontSize={30}
          onClick={() => onDelete()}
          additionalStyling="bg-emerald-300 hover:bg-green-400"
        />
        <Button
          label={'No'}
          width={75}
          fontSize={30}
          additionalStyling="bg-red-300 hover:bg-red-400"
          onClick={() => setAreYouSure(false)}
        />
      </div>
    </div>
  );
};

export default AreYouSure;
