import InputLabel from "@mui/material/InputLabel";
import MenuItem from "@mui/material/MenuItem";
import FormControl from "@mui/material/FormControl";
import Select from "@mui/material/Select";
import { useState } from "react";
import axios from "axios";
import Button from "@/components/Button";

const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

const PaginationInputs = ({ setDatabaseItems }) => {
  const [type, setType] = useState("book");
  const [limit, setLimit] = useState(5);
  const [sortBy, setSortBy] = useState("title");
  const [page, setPage] = useState(1);

  const handleGetMedia = async () => {
    try {
      const res = await axios.get(
        `${serverDomain}/database?type=${type}&sort=${sortBy}&limit=${limit}&page=${page}`,
      );
      const databaseResults = res.data;
      setDatabaseItems({ type, items: databaseResults.paginatedList });
    } catch (err) {
      console.log(`Server Error getting database items:`, err);
    }
  };

  let sortOptions;
  if (type === "book") {
    sortOptions = [
      { label: "Title", value: "title" },
      { label: "Author", value: "author" },
      { label: "Page Count", value: "page_count" },
      { label: "Pub. Year", value: "pub_year" },
    ];
  } else {
    sortOptions = [{ label: "Title", value: "title" }];
  }

  const handleTypeChange = (e) => {
    setType(e.target.value);
    if (e.target.value !== "book") {
      setSortBy("title");
    }
  };
  const handleLimitChange = (e) => {
    setLimit(e.target.value);
  };
  const handleSortByChange = (e) => {
    setSortBy(e.target.value);
  };
  return (
    <div className="border-3 w-3/10 mt-6 flex items-center justify-between rounded-lg border-[var(--darkpink)] bg-[var(--lightpink)] p-2 shadow-[5px_5px_30px_rgba(0,0,0,0.3)]">
      <FormControl sx={{ m: 1, minWidth: 130 }} size="small">
        <InputLabel id="type">Media Type</InputLabel>
        <Select
          labelId="type"
          id="type"
          value={type}
          label="Media Type"
          onChange={handleTypeChange}
        >
          <MenuItem value={"book"}>Book</MenuItem>
          <MenuItem value={"movie"}>Movie</MenuItem>
          <MenuItem value={"video_game"}>Video Game</MenuItem>
          <MenuItem value={"album"}>Album</MenuItem>
        </Select>
      </FormControl>
      <FormControl sx={{ m: 1, minWidth: 80 }} size="small">
        <InputLabel id="limit">Limit</InputLabel>
        <Select
          labelId="limit"
          id="limit"
          value={limit}
          label="limit"
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
        label={"Get Media"}
        width={100}
        fontSize={25}
      />
    </div>
  );
};

export default PaginationInputs;
