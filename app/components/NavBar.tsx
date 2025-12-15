// components
import Image from 'next/image';
import Link from 'next/link';
import Button from '@/app/components/Button';
import ProfileButton from '@/app/components/ProfileButton';

// public image imports
import Logo from '@/public/Comfort.png';

// auth
import { auth } from '@/lib/auth';
import { headers } from 'next/headers';

export default async function NavBar() {
  const session = await auth.api.getSession({ headers: await headers() });

  return (
    <nav className="bg-lightpink border-b-darkpink relative z-10 flex h-20 items-center gap-4 border-b-5 p-1.25">
      <Link href={'/'}>
        <Image
          src={Logo}
          alt="Megs Comfort Creations Logo"
          width={65}
          className="rounded-sm"
        />
      </Link>

      <h2 className="text-4xl tracking-wider">
        Welcome to Meg&apos;s Comfort Creations!
      </h2>
      <div className="ml-auto flex h-full flex-row items-center gap-5 pr-5">
        <Link href={'/'}>
          <Button label="Home" width={180} fontSize={36} />
        </Link>
        {session?.user.role === 'admin' && (
          <Link href={'/mediacollector'}>
            <Button label="Media Collector" width={180} fontSize={36} />
          </Link>
        )}
        {session?.user.role === 'admin' && (
          <Link href={'/showdatabase'}>
            <Button label="Show Database" width={180} fontSize={36} />
          </Link>
        )}
        <ProfileButton />
      </div>
    </nav>
  );
}
