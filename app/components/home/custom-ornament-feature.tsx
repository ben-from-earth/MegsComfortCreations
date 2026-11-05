import Image from 'next/image';
import Link from 'next/link';
import {
  homeClipPaths,
  homeCustomOrnament,
} from '@/components/home/home-content';

export default function CustomOrnamentFeature() {
  return (
    <section
      aria-label={homeCustomOrnament.eyebrow}
      className="relative mx-3 mt-3 h-150 overflow-hidden bg-[#f6a27e] max-lg:flex max-lg:h-auto max-lg:flex-col max-lg:items-center max-lg:gap-10 max-lg:px-6 max-lg:pt-8 max-lg:pb-12"
    >
      <div
        className="absolute top-9 left-9 z-2 h-132 w-152.5 -rotate-2 shadow-[0_14px_30px_rgba(60,30,15,0.3)] max-lg:relative max-lg:inset-auto max-lg:aspect-610/528 max-lg:h-auto max-lg:w-full max-lg:max-w-152.5"
        style={{ clipPath: homeClipPaths.customPhoto }}
      >
        <Image
          src={homeCustomOrnament.tornPhotoSrc}
          alt={homeCustomOrnament.tornPhotoAlt}
          fill
          sizes="(min-width: 1024px) 610px, 100vw"
          className="object-cover"
        />
      </div>
      <div className="absolute top-22.5 left-172.5 w-140 max-lg:static max-lg:w-full max-lg:max-w-152.5">
        <svg
          width="36"
          height="36"
          viewBox="0 0 24 24"
          fill="none"
          stroke="#c4452a"
          strokeWidth="1.6"
          strokeLinecap="round"
          strokeLinejoin="round"
          aria-hidden="true"
        >
          <path d="M2 5c3-1 6-1 10 1v14c-4-2-7-2-10-1z" />
          <path d="M22 5c-3-1-6-1-10 1v14c4-2 7-2 10-1z" />
          <path d="M5 9h4M5 12h4M15 9h4M15 12h4" />
        </svg>
        <p className="mt-3.5 text-[13px] font-extrabold tracking-[2.6px] text-[#c4452a] uppercase">
          {homeCustomOrnament.eyebrow}
        </p>
        <h2 className="mt-2.5 font-(family-name:--font-dm-serif-display) text-[56px] leading-none tracking-[-0.6px] text-[#1f1b18] max-sm:text-[40px]">
          {homeCustomOrnament.headline}
        </h2>
        <p className="mt-5 text-[17px] leading-normal text-[#3a2f28]">
          {homeCustomOrnament.body}
        </p>
        <Link
          href={homeCustomOrnament.button.href}
          className="mt-7 inline-flex h-14 items-center gap-3 rounded-md bg-[#24493b] px-7 text-[13px] font-extrabold tracking-[2px] text-[#f5eee2] uppercase shadow-[0_4px_0_#17302a]"
        >
          {homeCustomOrnament.button.label}
          <svg
            width="18"
            height="18"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2.4"
            strokeLinecap="round"
            strokeLinejoin="round"
            aria-hidden="true"
          >
            <path d="M5 12h14M13 6l6 6-6 6" />
          </svg>
        </Link>
      </div>
      <figure className="absolute top-17.5 right-14 m-0 w-80 rotate-5 bg-[#fffdf8] px-4 pt-4 shadow-[0_16px_34px_rgba(60,30,15,0.28)] max-lg:static max-lg:w-60 max-lg:rotate-3 max-lg:px-3 max-lg:pt-3">
        <div className="relative h-75 max-lg:h-52">
          <Image
            src={homeCustomOrnament.polaroidSrc}
            alt={homeCustomOrnament.polaroidAlt}
            fill
            sizes="288px"
            className="object-cover"
          />
        </div>
        <figcaption className="flex h-16 items-center justify-center font-(family-name:--font-just-another-hand) text-[34px] text-[#1f1b18]">
          {homeCustomOrnament.polaroidCaption}
        </figcaption>
      </figure>
    </section>
  );
}
