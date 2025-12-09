'use client';

import { useEffect, useState, useContext } from 'react';
import { useRouter } from 'next/navigation';
import axios from 'axios';

import Button from '@/app/components/Button';
import GenreContext from '@/lib/context/GenreContext';
import type { SuccessfulUserResponse } from '../api/auth/me/route';

const ProfilePage = () => {
  const router = useRouter();
  const genres = useContext(GenreContext);

  const [user, setUser] = useState<SuccessfulUserResponse['user'] | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const fetchMe = async () => {
      try {
        const res = await axios.get<SuccessfulUserResponse>('/api/auth/me', {
          withCredentials: true,
        });

        if (!res.data.access) {
          // Not logged in → redirect
          router.replace('/');
          return;
        }

        setUser(res.data.user);
      } catch (err) {
        console.error(err);
        router.replace('/');
      } finally {
        setLoading(false);
      }
    };

    fetchMe();
  }, [router]);

  const handleLogout = async () => {
    await axios.get('/api/auth/logout', {
      withCredentials: true,
    });
  };

  if (loading || !user) {
    return (
      <div className='m-5 text-center font-["Just_Another_Hand"] tracking-wider'>
        <h1 className="text-4xl">Loading profile…</h1>
      </div>
    );
  }

  const { first_name, last_name } = user;

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
          await handleLogout();
          router.replace('/');
        }}
        label="Log out"
        fontSize={32}
        width={200}
      />
    </div>
  );
};

export default ProfilePage;
