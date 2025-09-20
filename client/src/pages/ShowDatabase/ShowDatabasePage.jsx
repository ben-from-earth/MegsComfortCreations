import { useEffect, useState } from 'react';
import DatabaseItemDisplay from './DatabaseItemDisplay';
import PaginationInputs from './PaginationInputs';
import axios from 'axios';
import { titleRearrange } from '@/pages/MediaCollector/helpers/mediaCollectorHelpers';
import DatabasePageContext from '@/context/DatabasePageContext';

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

  useEffect(() => {
    handleGetMedia();
  }, [page, type, limit, sortBy, titleSearch]); //possible solution to constantly update list based on inputs
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
    <DatabasePageContext.Provider
      value={{
        page,
        setPage,
        databaseItems,
        type,
        setType,
        limit,
        setLimit,
        sortBy,
        setSortBy,
        setTitleSearch,
        handleGetMedia,
      }}
    >
      <div className="flex flex-col items-center">
        <PaginationInputs />
        <DatabaseItemDisplay />
      </div>
    </DatabasePageContext.Provider>
  );
};

export default ShowDatabasePage;
