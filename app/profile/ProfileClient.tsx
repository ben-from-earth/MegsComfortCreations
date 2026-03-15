'use client';

import { useContext, useMemo, useState } from 'react';
import { useRouter } from 'next/navigation';
import Button from '@/components/ui/Button';
import GenreContext from 'lib/context/GenreContext';
import { authClient } from 'lib/auth-client';
import { trpc } from 'lib/trpc/client';

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
  const [migrationStatusMessage, setMigrationStatusMessage] = useState<
    string | null
  >(null);

  const favoriteGenre = useMemo(() => {
    if (!genres?.length) return 'Unknown';
    return genres[Math.floor(Math.random() * genres.length)];
  }, [genres]);

  const booksRead = useMemo(() => Math.floor(Math.random() * 101), []);
  const imageMigrationStatus = trpc.profile.getImageMigrationStatus.useQuery();
  const migrateImageFilesMutation = trpc.profile.migrateImageFiles.useMutation();

  const migrationButtonDisabled =
    imageMigrationStatus.data?.isCompleted === true ||
    migrateImageFilesMutation.isPending;

  const handleImageMigration = async () => {
    setMigrationStatusMessage('Migration running...');
    try {
      const response = await migrateImageFilesMutation.mutateAsync();
      if (response.alreadyCompleted) {
        setMigrationStatusMessage(
          'Image migration already completed. No external image URLs remain.',
        );
      } else {
        const summary = response.summary;
        setMigrationStatusMessage(
          `Migration complete. Migrated ${summary?.migratedExternalUrls ?? 0} external images, ${summary?.failedDownloads ?? 0} failures, deleted ${summary?.deletedRows ?? 0} media rows.`,
        );
      }
      await imageMigrationStatus.refetch();
    } catch (error) {
      const message =
        error instanceof Error ? error.message : 'Image migration failed';
      setMigrationStatusMessage(`Migration failed: ${message}`);
    }
  };

  return (
    <div className='m-5 text-center font-["Just_Another_Hand"] tracking-wider'>
      <h1 className="mb-5 text-7xl">Hello, {user.name}!</h1>

      <h3 className="text-3xl">Favorite Book: [Book Title]</h3>
      <h3 className="text-3xl">Number of books read: {booksRead}</h3>
      <h3 className="mb-5 text-3xl">Favorite genre: {favoriteGenre}</h3>
      <div className="mb-5 flex flex-col items-center gap-2">
        <Button
          variant="primary"
          label={
            migrateImageFilesMutation.isPending
              ? 'Migrating image files...'
              : 'Migrate all image files'
          }
          onClick={handleImageMigration}
          disabled={migrationButtonDisabled}
          width={280}
          fontSize={28}
        />
        {imageMigrationStatus.data && (
          <>
            <p className="text-2xl">
              External image URLs remaining: {imageMigrationStatus.data.externalUrlCount}
            </p>
            <p className="text-2xl">
              Media items missing image records: {imageMigrationStatus.data.missingReferenceCount}
            </p>
          </>
        )}
        {migrationStatusMessage && (
          <p className="max-w-2xl text-2xl">{migrationStatusMessage}</p>
        )}
        <p className="max-w-2xl text-xl opacity-80">
          Temporary migration tool. Remove this button after rollout is confirmed.
        </p>
      </div>

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
