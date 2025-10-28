import { Suspense } from 'react';
import ProfileData from './ProfileData';

export default async function ProfilePage() {
  return (
    <div>
      <p>Profile Page</p>
      <Suspense fallback={<div>Loading profile data...</div>}>
        <ProfileData />
      </Suspense>
    </div>
  );
}
