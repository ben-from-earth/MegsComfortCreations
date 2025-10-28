// react, redux imports
import { useSelector } from 'react-redux';

// components
import Button from '@/app/components/Button';

//interfaces and types
import {
  databaseDataPerType,
  selectDatabaseData,
} from '@/lib/state/slices/databaseDataSlice';

interface ButtonGroupProps {
  onCollect: () => void;
  onPNG: () => Promise<void>;
  onDatabase: (databaseDate: databaseDataPerType[]) => Promise<void>;
}

export default function ButtonGroup({
  onCollect,
  onPNG,
  onDatabase,
}: ButtonGroupProps) {
  // setup connection to redux slice and get the database information
  const databaseData: databaseDataPerType[] = useSelector(selectDatabaseData);

  return (
    <div className="flex flex-row content-center gap-4">
      <Button
        onClick={() => {
          onCollect();
        }}
        label={'Collect Media Covers'}
        width={175}
        fontSize={25}
      />
      <Button
        onClick={() => onDatabase(databaseData)}
        label={'Send to Database'}
        width={175}
        fontSize={25}
      />
      <Button
        onClick={() => onPNG()}
        label={'Get PNG'}
        width={175}
        fontSize={25}
      />
    </div>
  );
}
