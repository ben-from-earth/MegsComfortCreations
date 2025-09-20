import InputLabel from '@mui/material/InputLabel';
import MenuItem from '@mui/material/MenuItem';
import FormControl from '@mui/material/FormControl';
import Select from '@mui/material/Select';
import Button from '@/components/Button';
import { useContext } from 'react';
import DatabasePageContext from '@/context/DatabasePageContext';

const PaginationInputs = () => {
  const {
    type,
    setType,
    limit,
    setLimit,
    sortBy,
    setSortBy,
    page,
    setPage,
    setTitleSearch,
    handleGetMedia,
  } = useContext(DatabasePageContext);
  let sortOptions;
  if (type === 'book') {
    sortOptions = [
      { label: 'Title', value: 'title' },
      { label: 'Author', value: 'author' },
      { label: 'Page Count', value: 'page_count' },
      { label: 'Pub. Year', value: 'pub_year' },
    ];
  } else {
    sortOptions = [{ label: 'Title', value: 'title' }];
  }

  const handleTypeChange = (e) => {
    setPage(1);
    setType(e.target.value);

    if (e.target.value !== 'book') {
      setSortBy('title');
    }
  };
  const handleLimitChange = (e) => {
    setPage(1);
    setLimit(e.target.value);
  };
  const handleSortByChange = (e) => {
    setPage(1);
    setSortBy(e.target.value);
  };
  return (
    <div className="border-3 mt-6 flex w-fit items-center justify-between rounded-lg border-[var(--darkpink)] bg-[var(--lightpink)] p-2 shadow-[5px_5px_30px_rgba(0,0,0,0.3)]">
      <input
        id="titleSearch"
        onChange={(e) => setTitleSearch(e.target.value)}
        placeholder="Title"
        className="h-10 rounded-sm border border-[rgba(0,0,0,0.23)] bg-[var(--lightpink)] pl-2"
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
          <MenuItem value={'video_game'}>Video Game</MenuItem>
          <MenuItem value={'album'}>Album</MenuItem>
        </Select>
      </FormControl>
      <FormControl sx={{ m: 1, minWidth: 80 }} size="small">
        <InputLabel id="limit">Per Page</InputLabel>
        <Select
          labelId="limit"
          id="limit"
          value={limit}
          label="Per Page"
          onChange={handleLimitChange}
        >
          <MenuItem value={3}>3</MenuItem>
          <MenuItem value={5}>5</MenuItem>
          <MenuItem value={10}>10</MenuItem>
        </Select>
      </FormControl>
      <FormControl sx={{ m: 1, minWidth: 100 }} size="small">
        <InputLabel id="sortBy">Sort By</InputLabel>
        <Select
          labelId="sortBy"
          id="sortBy"
          value={sortBy}
          label="sortBy"
          onChange={handleSortByChange}
        >
          {sortOptions.map((option) => (
            <MenuItem value={option.value}>{option.label}</MenuItem>
          ))}
        </Select>
      </FormControl>
      <Button
        onClick={handleGetMedia}
        label={'Search'}
        width={100}
        fontSize={25}
      />
    </div>
  );
};

export default PaginationInputs;
