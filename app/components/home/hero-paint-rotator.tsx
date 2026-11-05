import { existsSync } from 'node:fs';
import path from 'node:path';
import Image from 'next/image';
import Link from 'next/link';
import {
  homeClipPaths,
  homeHeroCopy,
  homeWipes,
} from '@/components/home/home-content';
import './home.css';

const publicDirectory = path.join(process.cwd(), 'public');

export default function HeroPaintRotator() {
  return (
    <section className="relative h-160 overflow-hidden bg-[#3b2a22]">
      <div aria-hidden="true" className="absolute inset-[0_0_0_520px]">
        {homeWipes.map((wipe, index) => (
          <div
            key={wipe.word}
            className="home-hero-photo"
            style={{ animationDelay: `${wipe.delaySeconds}s` }}
          >
            {existsSync(path.join(publicDirectory, wipe.photoSrc)) ? (
              <Image
                src={wipe.photoSrc}
                alt={wipe.photoAlt}
                fill
                sizes="(min-width: 1024px) calc(100vw - 520px), 100vw"
                priority={index === 0}
                className="object-cover"
              />
            ) : null}
            {wipe.tag ? (
              <div
                className="absolute top-20 right-17.5 z-3 flex w-37.5 -rotate-6 justify-center bg-[#f2e3c8] px-2.5 py-4.5 shadow-[0_8px_18px_rgba(0,0,0,0.3)]"
                style={{ clipPath: homeClipPaths.photoTag }}
              >
                <span className="text-center font-(family-name:--font-just-another-hand) text-[40px] leading-[0.85] text-[#1f1b18] uppercase">
                  {wipe.tag[0]}
                  <br />
                  {wipe.tag[1]}
                </span>
              </div>
            ) : null}
          </div>
        ))}
      </div>

      <div
        className="absolute inset-y-0 left-0 z-2 w-225 bg-[#f6efe3]"
        style={{ clipPath: homeClipPaths.heroPanel }}
      >
        <div className="absolute top-25 left-32.5 w-155">
          <h1 className="m-0 flex flex-col items-start">
            <span className="font-(family-name:--font-dm-serif-display) text-[96px] leading-[0.95] tracking-[-1.5px] text-[#1f1b18]">
              {homeHeroCopy.headline}
            </span>
            <span className="sr-only">{homeHeroCopy.hiddenTitle}</span>
            <span
              aria-hidden="true"
              className="pointer-events-none relative mt-2.5 ml-5 block h-50 w-130"
            >
              {homeWipes.map((wipe, index) => (
                <span
                  key={wipe.word}
                  className="home-wipe"
                  style={{ animationDelay: `${wipe.delaySeconds}s` }}
                >
                  <Image
                    src={wipe.src}
                    alt=""
                    fill
                    sizes="520px"
                    priority={index === 0}
                    className="object-cover object-[50%_42%]"
                  />
                </span>
              ))}
            </span>
          </h1>
          <p className="mt-6 text-2xl font-semibold text-[#1f1b18]">
            {homeHeroCopy.tagline}
          </p>
          <p className="mt-2 max-w-95 text-xl leading-[1.45] text-[#4a433d]">
            {homeHeroCopy.body}
          </p>
          <Link
            href={homeHeroCopy.button.href}
            className="mt-7.5 inline-flex h-14 items-center gap-3 rounded-md bg-[#24493b] px-7.5 text-sm font-extrabold tracking-[2.2px] text-[#f5eee2] uppercase shadow-[0_4px_0_#17302a]"
          >
            {homeHeroCopy.button.label}
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
      </div>
    </section>
  );
}
