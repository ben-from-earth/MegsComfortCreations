import { useDispatch } from 'react-redux';
import { updateDatabaseData } from '@/state/databaseDataSlice';

// setup component text area for each data field in the block
const MyTextArea = ({ name, label, type, blockID, value }) => {
  const dispatch = useDispatch();

  // const labelClass =
  //   type === 'book'
  //     ? 'absolute right-full w-fit border border-red-500 mr-2 top-1/2 -translate-y-1/2 font-["Just_Another_Hand"] text-3xl'
  //     : 'w-15 content-center text-right font-["Just_Another_Hand"] text-3xl';

  return (
    <div className="relative">
      <label
        className='translate-y-1/8 absolute right-full mr-2 w-fit text-nowrap text-right font-["Just_Another_Hand"] text-3xl'
        htmlFor={name}
      >
        {label}:
      </label>
      <textarea
        className="content-center rounded-sm bg-white pl-2 text-black"
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
};

export default MyTextArea;
