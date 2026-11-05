import { existsSync } from 'node:fs';
import path from 'node:path';
import type { CSSProperties } from 'react';
import Image from 'next/image';
import Link from 'next/link';
import {
  homeClipPaths,
  homeHeroCopy,
  homeWipes,
} from '@/components/home/home-content';
import './home.css';

const publicDirectory = path.join(process.cwd(), 'public');
const wordsWithHeroPhotos = new Set(
  homeWipes.flatMap((wipe) =>
    existsSync(path.join(publicDirectory, wipe.photoSrc)) ? [wipe.word] : [],
  ),
);

// The narrow top edge reuses the desktop tear's jitter with its slant removed.
// 1% of the 900px desktop panel is 9px.
const desktopTearPoints = Array.from(
  homeClipPaths.heroPanel.matchAll(/([\d.]+)% ([\d.]+)%/g),
  (match) => ({ x: Number(match[1]), y: Number(match[2]) }),
);
const firstTearPoint = desktopTearPoints[0];
const lastTearPoint = desktopTearPoints[desktopTearPoints.length - 1];
const narrowPanelClipPath = `polygon(${desktopTearPoints
  .map((point) => {
    const slantX =
      firstTearPoint.x + ((lastTearPoint.x - firstTearPoint.x) * point.y) / 100;
    return `${point.y}% ${(12 + (point.x - slantX) * 9).toFixed(1)}px`;
  })
  .join(',')},100% 100%,0 100%)`;

const heroPanelStyle: CSSProperties &
  Record<'--hero-panel-clip' | '--hero-panel-clip-narrow', string> = {
  '--hero-panel-clip': homeClipPaths.heroPanel,
  '--hero-panel-clip-narrow': narrowPanelClipPath,
};

export default function HeroPaintRotator() {
  return (
    <section className="relative h-160 overflow-hidden bg-[#3b2a22] max-lg:flex max-lg:h-auto max-lg:flex-col">
      <div
        aria-hidden="true"
        className="absolute inset-[0_0_0_520px] max-lg:relative max-lg:inset-auto max-lg:h-90 max-lg:w-full"
      >
        {homeWipes.map((wipe, index) => (
          <div
            key={wipe.word}
            className="home-hero-photo"
            style={{ animationDelay: `${wipe.delaySeconds}s` }}
          >
            {wordsWithHeroPhotos.has(wipe.word) ? (
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
        className="home-hero-panel absolute inset-y-0 left-0 z-2 w-225 bg-[#ffedd9] max-lg:relative max-lg:inset-auto max-lg:-mt-6 max-lg:w-full max-lg:bg-[#ffedd9] max-lg:pb-14 max-lg:[clip-path:var(--hero-panel-clip-narrow)]"
        style={heroPanelStyle}
      >
        <div className="pointer-events-none absolute inset-0 z-3">
          <Image
            src="/home/hero-paint-smears.png"
            alt=""
            fill
            priority
            unoptimized
            sizes="900px"
            className="object-fill"
          />
        </div>
        <div className="absolute top-22 left-40 z-2 w-155 max-lg:static max-lg:w-auto max-lg:px-6 max-lg:pt-14">
          <h1 className="m-0 flex flex-col items-start">
            <span className="font-(family-name:--font-dm-serif-display) text-[96px] leading-[0.95] tracking-[-1.5px] text-[#1f1b18] max-sm:text-[64px]">
              {homeHeroCopy.headline}
            </span>
            <span className="sr-only">{homeHeroCopy.hiddenTitle}</span>
            <span
              aria-hidden="true"
              className="pointer-events-none relative mt-2.5 ml-5 block h-50 w-130 origin-top-left max-[400px]:-mb-24 max-[400px]:scale-52 max-sm:-mb-20 max-sm:ml-0 max-sm:scale-60"
            >
              {homeWipes.map((wipe, index) => (
                <span
                  key={wipe.word}
                  className="home-wipe"
                  style={{
                    animationDelay: `${wipe.delaySeconds}s`,
                    transform: `rotate(${wipe.tiltDegrees}deg)`,
                  }}
                >
                  <Image
                    src={wipe.src}
                    alt=""
                    fill
                    sizes="520px"
                    priority={index === 0}
                    unoptimized
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
            {homeHeroCopy.body}{' '}
            <span className="font-(family-name:--font-just-another-hand) text-[32px] leading-none text-[#1f1b18]">
              {homeHeroCopy.signature}
            </span>
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
