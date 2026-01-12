import { auth } from 'lib/auth';
import { headers } from 'next/headers';
import { redirect } from 'next/navigation';
import ProfileClient from './ProfileClient';

export default async function ProfilePage() {
  const session = await auth.api.getSession({ headers: await headers() });

  if (!session) redirect('/login');

  // session.user contains your user info
  return <ProfileClient user={session.user} />;
}
