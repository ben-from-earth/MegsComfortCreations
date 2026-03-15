'use client';

import { useContext, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import Button from '@/components/ui/Button';
import GenreContext from 'lib/context/GenreContext';
import { authClient } from 'lib/auth-client';

type Props = {
  user: {
    // keep flexible; Better Auth user shape depends on your schema
    name?: string | null;
    email?: string | null;
  };
};

const ProfileClient = ({ user }: Props) => {
  const router = useRouter();
  const genres = useContext(GenreContext);

  const favoriteGenre = useMemo(() => {
    if (!genres?.length) return 'Unknown';
    return genres[Math.floor(Math.random() * genres.length)];
  }, [genres]);

  const booksRead = useMemo(() => Math.floor(Math.random() * 101), []);

  return (
    <div className='m-5 text-center font-["Just_Another_Hand"] tracking-wider'>
      <h1 className="mb-5 text-7xl">Hello, {user.name}!</h1>

      <h3 className="text-3xl">Favorite Book: [Book Title]</h3>
      <h3 className="text-3xl">Number of books read: {booksRead}</h3>
      <h3 className="mb-5 text-3xl">Favorite genre: {favoriteGenre}</h3>

      <Button
        variant="primary"
        onClick={async () => {
          await authClient.signOut();
          router.replace('/');
          router.refresh();
        }}
        label="Log out"
        fontSize={32}
        width={200}
      />
    </div>
  );
};

export default ProfileClient;
