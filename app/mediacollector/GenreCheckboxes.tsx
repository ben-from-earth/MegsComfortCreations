// react imports
import { ChangeEvent } from 'react';

// interfaces and types

export interface GenreCheckboxProps {
  genres: string[];
  databaseGenres: string[];
  handleGenreClick: (genreText: string, checked: boolean) => void;
}

export default function GenreCheckboxes({
  genres,
  databaseGenres,
  handleGenreClick,
}: GenreCheckboxProps) {
  return (
    <div className="mx-6.25 mb-7.5 grid grid-cols-[repeat(4,max-content)] items-center justify-between gap-x-2.5 gap-y-0.75">
      {genres?.map((text, idx) => (
        <label key={idx} className="cursor-pointer text-xl">
          <input
            type="checkbox"
            className="m-1"
            defaultChecked={databaseGenres?.includes(text)}
            onChange={(e: ChangeEvent<HTMLInputElement>) => {
              handleGenreClick(text, e.target.checked);
            }}
          />
          {text}
        </label>
      ))}
    </div>
  );
}
