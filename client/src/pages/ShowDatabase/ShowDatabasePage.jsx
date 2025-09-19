import { useState } from 'react';
import DatabaseItemDisplay from './DatabaseItemDisplay';
import PaginationInputs from './PaginationInputs';
import axios from 'axios';
import { titleRearrange } from '../MediaCollector/helpers/mediaCollectorHelpers';

const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

const ShowDatabasePage = () => {
  const [databaseItems, setDatabaseItems] = useState({
    type: '',
    items: [],
    total: 0,
    min: 0,
    max: 0,
  });

  const [type, setType] = useState('book');
  const [limit, setLimit] = useState(5);
  const [sortBy, setSortBy] = useState('title');
  const [page, setPage] = useState(1);
  const [titleSearch, setTitleSearch] = useState('');

  const handleGetMedia = async () => {
    if (titleSearch.length > 0) {
      try {
        const res = await axios.get(
          `${serverDomain}/database/titleSearch?type=${type}&title=${titleRearrange(titleSearch)}`,
        );
        const databaseResults = res.data;
        setDatabaseItems({
          type,
          items: databaseResults.titleSearchResponse,
          total: databaseResults.total,
          min: 1,
          max: databaseResults.total,
        });
      } catch (err) {
        console.log(`Server Error getting database items:`, err);
      }
    } else {
      try {
        const res = await axios.get(
          `${serverDomain}/database?type=${type}&sort=${sortBy}&limit=${limit}&page=${page}`,
        );
        const databaseResults = res.data;
        setDatabaseItems({
          type,
          items: databaseResults.paginatedList,
          total: databaseResults.total,
          min: (page - 1) * limit + 1,
          max: page * limit,
        });
      } catch (err) {
        console.log(`Server Error getting database items:`, err);
      }
    }
  };
  return (
    <div className="flex flex-col items-center">
      <PaginationInputs
        paginationFunctions={{
          setType,
          setLimit,
          setSortBy,
          setPage,
          setTitleSearch,
          handleGetMedia,
        }}
        paginationVariables={{ type, limit, sortBy, page }}
      />
      <DatabaseItemDisplay
        databaseItems={databaseItems}
        handleGetMedia={handleGetMedia}
      />
    </div>
  );
};

export default ShowDatabasePage;
