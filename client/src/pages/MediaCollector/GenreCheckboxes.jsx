const GenreCheckboxes = ({ genres, databaseGenres, handleGenreClick }) => {
  return (
    <div className="gap-y-1.25 mx-6.25 mb-7.5 grid grid-cols-[repeat(4,max-content)] items-center justify-between gap-x-2.5 text-2xl">
      {genres?.map((text, idx) => (
        <label
          key={idx}
          name={text}
          className='cursor-pointer font-["Just_Another_Hand"]'
          onChange={(e) => {
            handleGenreClick(text, e.target.checked);
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
