import Image from 'next/image';
import Link from 'next/link';
import { DM_Serif_Display } from 'next/font/google';
import NavLinks from '@/components/shared/nav-links';
import ProfileButton from '@/components/shared/profile-button';
import { auth } from 'lib/auth';
import { headers } from 'next/headers';

const dmSerifDisplay = DM_Serif_Display({ weight: '400', subsets: ['latin'] });

export default async function NavBar() {
  const session = await auth.api.getSession({ headers: await headers() });

  return (
    <header className="relative z-10 bg-[#f3f6f0] text-[#2b2826] shadow-[0_2px_10px_rgba(60,50,40,.06)]">
      <div className="flex h-[84px] items-center gap-14 px-10">
        <Link
          href="/"
          aria-label="Comfort Creations by Meg, home"
          className="flex flex-none items-center gap-2.5 text-[#2b2826] no-underline"
        >
          <Image
            src="/logo-mark.png"
            alt=""
            aria-hidden="true"
            width={160}
            height={146}
            className="block h-[58px] w-auto"
            priority
          />
          <span className="flex flex-col leading-none">
            <span
              className={`${dmSerifDisplay.className} text-[30px] tracking-[-0.3px]`}
            >
              Comfort Creations
            </span>
            <span className="mt-px flex items-center gap-2">
              <span
                aria-hidden="true"
                className="h-[1.5px] flex-1 rounded-[1px] bg-[#2b2826]"
              />
              <span className="text-[26px] leading-[0.9]">by Meg</span>
            </span>
          </span>
        </Link>

        <NavLinks isAdmin={session?.user.role === 'admin'} />

        <ProfileButton />
      </div>
      <div
        aria-hidden="true"
        className="h-1.5 bg-[linear-gradient(90deg,#b8c4a7_0_18%,#f0c9a8_18%_31%,#e9d9c4_31%_52%,#d9b2b0_52%_63%,#b8c4a7_63%_81%,#f0c9a8_81%_100%)] opacity-85"
      />
    </header>
  );
}
