//react imports
import { useDispatch } from "react-redux";

//imports from database state slice
import {
  addToDatabaseData,
  removeFromDatabaseData,
} from "@/state/databaseDataSlice";

const GenreCheckboxes = ({ genres, databaseGenres, blockID }) => {
  const dispatch = useDispatch();

  //if genre is clicked we add it to the data associated with the block and remove if unchecked
  const handleGenreClick = (genreText, checked, blockID) => {
    const type = "book";
    if (checked) {
      dispatch(addToDatabaseData({ blockID, type, genreText }));
    } else {
      dispatch(removeFromDatabaseData({ blockID, type, genreText }));
    }
  };

  return (
    <div className="gap-y-1.25 mx-6.25 mb-7.5 grid grid-cols-[repeat(4,max-content)] items-center justify-between gap-x-2.5 text-2xl">
      {genres?.map((text, idx) => (
        <label
          key={idx}
          name={text}
          className='cursor-pointer font-["Just_Another_Hand"]'
          onChange={(e) => {
            handleGenreClick(text, e.target.checked, blockID);
          }}
        >
          <input
            type="checkbox"
            className="m-1"
            defaultChecked={databaseGenres?.includes(text)}
          />
          {text}
        </label>
      ))}
    </div>
  );
};

export default GenreCheckboxes;
