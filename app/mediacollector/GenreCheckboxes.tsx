// react imports
import { Combobox } from '@/components/shared/combobox';
import Button from '@/components/ui/Button';

// interfaces and types

export interface GenreCheckboxProps {
  allGenres: string[];
  handleGenreClick: (genreText: string, checked: boolean) => void;
  index?: number;
  blockGenres?: string[];
}

export default function GenreCheckboxes({
  allGenres,
  handleGenreClick,
  blockGenres,
}: GenreCheckboxProps) {
  const genres = blockGenres || [];

  // database edit page

  return (
    <div className="mx-6 mb-7.5 flex w-full flex-col px-5">
      <Combobox
        items={allGenres
          .filter((genre) => !genres.includes(genre))
          .map((genre) => ({ value: genre }))}
        onSelect={(value) => {
          handleGenreClick(value, true);
        }}
        label="genres"
      />
      <div className="mt-2 flex flex-wrap">
        {genres.map((genre) => (
          <GenreTag
            key={genre}
            genre={genre}
            onClick={() => handleGenreClick(genre, false)}
          />
        ))}
      </div>
    </div>
  );
}

const GenreTag = ({
  genre,
  onClick,
}: {
  genre: string;
  onClick: () => void;
}) => {
  return (
    <div className="border-darkpink bg-lightpink m-0.5 flex h-8 items-center gap-1 rounded border-2 px-1 text-lg text-black shadow-[0px_2px_6px_rgba(0,0,0,0.3)]">
      <Button variant="remove" onClick={onClick} />
      <span>{genre}</span>
    </div>
  );
};
