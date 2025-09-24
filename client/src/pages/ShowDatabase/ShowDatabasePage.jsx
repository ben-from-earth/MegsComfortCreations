//react
import { useEffect, useState } from 'react';

//components
import DatabaseItemDisplay from '@/pages/ShowDatabase/DatabaseItemDisplay';
import PaginationInputs from '@/pages/ShowDatabase/PaginationInputs';

//axios
import axios from 'axios';

//helpers
import { titleRearrange } from '@/pages/MediaCollector/helpers/mediaCollectorHelpers';

//context
import DatabasePageContext from '@/context/DatabasePageContext';

//server domain for axios requests
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
  const [genre, setGenre] = useState('');
  const [ascDesc, setAscDesc] = useState('asc');

  useEffect(() => {
    handleGetMedia();
  }, [page, type, limit, sortBy, titleSearch, genre, ascDesc]);
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
      if (type !== 'book' || genre === '') {
        try {
          const res = await axios.get(
            `${serverDomain}/database?type=${type}&sort=${sortBy}&limit=${limit}&page=${page}&ascDesc=${ascDesc}`,
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
      } else {
        if (genre === 'none') {
          const res = await axios.get(
            `${serverDomain}/genres/nogenres?sort=${sortBy}&limit=${limit}&page=${page}&ascDesc=${ascDesc}`,
          );
          const databaseResults = res.data;

          setDatabaseItems({
            type,
            items: databaseResults.paginatedList,
            total: databaseResults.total,
            min: (page - 1) * limit + 1,
            max: page * limit,
          });
        } else {
          const res = await axios.get(
            `${serverDomain}/genres?genre=${genre}&sort=${sortBy}&limit=${limit}&page=${page}&ascDesc=${ascDesc}`,
          );
          const databaseResults = res.data;

          setDatabaseItems({
            type,
            items: databaseResults.paginatedList,
            total: databaseResults.total,
            min: (page - 1) * limit + 1,
            max: page * limit,
          });
        }
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
        genre,
        setGenre,
        ascDesc,
        setAscDesc,
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
