// react, redux imports
import { useContext } from 'react';

// import icons and items from Material UI
import InputLabel from '@mui/material/InputLabel';
import MenuItem from '@mui/material/MenuItem';
import FormControl from '@mui/material/FormControl';
import Select, { SelectChangeEvent } from '@mui/material/Select';
import Button from '@/components/ui/Button';

// context
import GenreContext from 'lib/context/GenreContext';

// interfaces and types

import {
  DatabasePageContextValue,
  useDatabasePageContext,
} from 'lib/context/DatabasePageContext';

import { MediaType } from 'lib/interfaces/globalInterfaces';
import { NO_GENRE_FILTER } from '@/lib/enums/genreEnums';

export type SortOptions = 'title' | 'author' | 'pageCount' | 'pubYear';
export type SortOptionsLabels = 'Title' | 'Author' | 'Page Count' | 'Pub. Year';

export default function PaginationInputs() {
  const {
    type,
    setType,
    limit,
    setLimit,
    sortBy,
    setSortBy,
    setPage,
    genre,
    setGenre,
    ascDesc,
    setAscDesc,
    setTitleSearch,
    handleGetMedia,
  }: DatabasePageContextValue = useDatabasePageContext();
  const genres = useContext(GenreContext);
  let sortOptions: { label: SortOptionsLabels; value: SortOptions }[];
  if (type === 'book') {
    sortOptions = [
      { label: 'Title', value: 'title' },
      { label: 'Author', value: 'author' },
      { label: 'Page Count', value: 'pageCount' },
      { label: 'Pub. Year', value: 'pubYear' },
    ];
  } else {
    sortOptions = [{ label: 'Title', value: 'title' }];
  }
  const handleTypeChange = (e: SelectChangeEvent) => {
    setPage(1);
    setType(e.target.value as MediaType);

    if (e.target.value !== 'book') {
      setSortBy('title');
      setGenre('');
    }
  };
  const handleLimitChange = (e: SelectChangeEvent) => {
    setPage(1);
    setLimit(Number(e.target.value) as 3 | 5 | 10);
  };
  const handleSortByChange = (e: SelectChangeEvent) => {
    setPage(1);
    setSortBy(e.target.value as SortOptions);
  };
  const handleGenreChange = (e: SelectChangeEvent) => {
    setPage(1);
    setGenre(
      e.target.value as (typeof genres)[number] | '' | typeof NO_GENRE_FILTER,
    );
  };
  const handleAscDescChange = (e: SelectChangeEvent) => {
    setPage(1);
    setAscDesc(e.target.value as 'asc' | 'desc');
  };
  return (
    <div className="border-darkpink bg-lightpink mt-6 flex w-fit items-center justify-between rounded-lg border-3 p-2 shadow-[5px_5px_30px_rgba(0,0,0,0.3)]">
      <input
        id="titleSearch"
        onChange={(e) => setTitleSearch(e.target.value)}
        placeholder="Title"
        className="bg-lightpink w-xxs h-10 rounded-sm border border-[rgba(0,0,0,0.23)] pl-2 font-[Arial] text-sm"
      ></input>
      <FormControl sx={{ m: 1, minWidth: 130 }} size="small">
        <InputLabel id="type">Media Type</InputLabel>
        <Select
          labelId="type"
          id="type"
          value={type}
          label="Media Type"
          onChange={handleTypeChange}
        >
          <MenuItem value={'book'}>Book</MenuItem>
          <MenuItem value={'movie'}>Movie</MenuItem>
          <MenuItem value={'videoGame'}>Video Game</MenuItem>
          <MenuItem value={'album'}>Album</MenuItem>
        </Select>
      </FormControl>
      <FormControl sx={{ m: 1, minWidth: 80 }} size="small">
        <InputLabel id="limit">Per Page</InputLabel>
        <Select
          labelId="limit"
          id="limit"
          value={String(limit)}
          label="Per Page"
          onChange={handleLimitChange}
        >
          <MenuItem value={3}>3</MenuItem>
          <MenuItem value={5}>5</MenuItem>
          <MenuItem value={10}>10</MenuItem>
        </Select>
      </FormControl>
      <FormControl
        sx={{ m: 1, minWidth: 100 }}
        size="small"
        disabled={type !== 'book'}
      >
        <InputLabel id="sortBy">Sort By</InputLabel>
        <Select
          labelId="sortBy"
          id="sortBy"
          value={sortBy}
          label="sortBy"
          onChange={handleSortByChange}
        >
          {sortOptions.map((option) => (
            <MenuItem key={option.label} value={option.value}>
              {option.label}
            </MenuItem>
          ))}
        </Select>
      </FormControl>
      <FormControl
        sx={{ m: 1, minWidth: 100 }}
        size="small"
        disabled={type !== 'book'}
      >
        <InputLabel id="genre">Genre</InputLabel>
        <Select
          labelId="genre"
          id="genre"
          value={genre}
          label="genre"
          onChange={handleGenreChange}
        >
          <MenuItem value={''}>---</MenuItem>
          <MenuItem value={NO_GENRE_FILTER}>None</MenuItem>

          {genres.map((option) => (
            <MenuItem key={option} value={option}>
              {option}
            </MenuItem>
          ))}
        </Select>
      </FormControl>
      <FormControl sx={{ m: 1, minWidth: 100 }} size="small">
        <InputLabel id="ascDesc">Asc/Desc</InputLabel>
        <Select
          labelId="ascDesc"
          id="ascDesc"
          value={ascDesc}
          label="ascDesc"
          onChange={handleAscDescChange}
        >
          <MenuItem value={'asc'}>Ascending</MenuItem>
          <MenuItem value={'desc'}>Descending</MenuItem>
        </Select>
      </FormControl>
      <Button
        variant="primary"
        onClick={handleGetMedia}
        label={'Search'}
        width={100}
        fontSize={25}
      />
    </div>
  );
}
