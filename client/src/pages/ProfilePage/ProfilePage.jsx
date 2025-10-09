import { useLoaderData, useNavigate } from 'react-router';
import Button from '@/components/Button';
import axios from 'axios';

//server location import from .env
const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

//genres from context provider to populate genre list based on what genres are in the database
import GenreContext from '@/context/GenreContext';
import { useContext } from 'react';

const ProfilePage = () => {
  const navigate = useNavigate();
  const user = useLoaderData();

  //get genres for checkbox population
  const genres = useContext(GenreContext);

  const { first_name, last_name } = user;

  const handleLogout = async () => {
    const res = await axios.get(`${serverDomain}/auth/logout`, {
      withCredentials: true,
    });
  };
  return (
    <div className='m-5 text-center font-["Just_Another_Hand"] tracking-wider'>
      <h1 className="mb-5 text-7xl">
        Hello, {first_name} {last_name}!
      </h1>
      <h3 className="text-3xl">Favorite Book: [Book Title]</h3>
      <h3 className="text-3xl">
        Number of books read: {Math.floor(Math.random() * 101)}
      </h3>
      <h3 className="mb-5 text-3xl">
        Favorite genre: {genres[Math.floor(Math.random() * genres.length)]}
      </h3>
      <Button
        onClick={async () => {
          handleLogout();
          navigate('/', { replace: true });
        }}
        label="Log out"
        fontSize={32}
      />
    </div>
  );
};

export default ProfilePage;
