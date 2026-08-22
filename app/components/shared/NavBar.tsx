// components
import Image from 'next/image';
import Link from 'next/link';
import Button from '@/components/ui/Button';
import ProfileButton from '@/components/shared/ProfileButton';

// auth
import { auth } from 'lib/auth';
import { headers } from 'next/headers';

const logoImagePath = '/Comfort.png';

export default async function NavBar() {
  const session = await auth.api.getSession({ headers: await headers() });

  return (
    <nav className="bg-lightpink border-b-darkpink relative z-10 flex h-20 items-center gap-4 border-b-5 p-1.25">
      <Link href={'/'}>
        <Image
          src={logoImagePath}
          alt="Megs Comfort Creations Logo"
          width={65}
          height={65}
          className="rounded-sm"
        />
      </Link>

      <h2 className="text-4xl tracking-wider">
        Welcome to Meg&apos;s Comfort Creations!
      </h2>
      <div className="ml-auto flex h-full flex-row items-center gap-5 pr-5">
        <Link href={'/'}>
          <Button variant="primary" label="Home" width={180} fontSize={36} />
        </Link>
        {session?.user.role === 'admin' && (
          <Link href={'/mediacollector'}>
            <Button
              variant="primary"
              label="Media Collector"
              width={180}
              fontSize={36}
            />
          </Link>
        )}
        {session?.user.role === 'admin' && (
          <Link href={'/showdatabase'}>
            <Button
              variant="primary"
              label="Show Database"
              width={180}
              fontSize={36}
            />
          </Link>
        )}
        <ProfileButton />
      </div>
    </nav>
  );
}
