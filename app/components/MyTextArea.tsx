// react, redux imports
import { useAppDispatch } from '@/lib/state/store';

// necessary imports from database data slice
import { updateDatabaseData } from '@/lib/state/slices/databaseDataSlice';

// interfaces and types
import { MediaType } from '@/lib/interfaces/globalInterfaces';

export interface MyTextAreaProps {
  name: 'title' | 'author' | 'pubYear' | 'pageCount';
  label: string;
  type: MediaType;
  blockID: string;
  value: string | number;
}

export default function MyTextArea({
  name,
  label,
  type,
  blockID,
  value,
}: MyTextAreaProps) {
  const dispatch = useAppDispatch();

  return (
    <div className="relative">
      <label
        className="absolute right-full mr-2 w-fit translate-y-1/8 text-right text-3xl text-nowrap"
        htmlFor={name}
      >
        {label}:
      </label>
      <textarea
        className="content-center rounded-sm bg-white pl-2 font-[Arial] text-black"
        style={
          type !== 'book'
            ? { marginBottom: '20px', width: '200px' }
            : { width: '300px' }
        }
        name={name}
        defaultValue={value}
        onChange={(e) => {
          dispatch(
            updateDatabaseData({
              blockID,
              type,
              name,
              newText: e.target.value,
            }),
          );
        }}
      ></textarea>
    </div>
  );
}
