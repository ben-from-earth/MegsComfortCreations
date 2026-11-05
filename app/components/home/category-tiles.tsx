import Image from 'next/image';
import Link from 'next/link';
import { homeClipPaths, homeTiles } from '@/components/home/home-content';

export default function CategoryTiles() {
  return (
    <section
      aria-label="Shop categories"
      className="grid grid-cols-5 gap-3 px-3 pt-3 max-lg:grid-cols-2"
    >
      {homeTiles.map((tile, index) => (
        <Link
          key={tile.href}
          href={tile.href}
          className="relative block h-75 overflow-hidden max-lg:last:col-span-2"
        >
          <Image
            src={tile.paperSrc}
            alt=""
            fill
            sizes="(min-width: 1024px) 20vw, 100vw"
            className="object-cover"
          />
          <div
            className="absolute top-6.5 left-1/2 -ml-23.75 h-42.5 w-47.5 overflow-hidden rounded-[14px] shadow-[0_10px_22px_rgba(0,0,0,0.25)]"
            style={{ transform: `rotate(${tile.photoTiltDegrees}deg)` }}
          >
            <Image
              src={tile.imageSrc}
              alt={tile.imageAlt}
              fill
              sizes="190px"
              className="object-cover"
            />
          </div>
          <span
            className="absolute right-5.5 bottom-5.5 left-5.5 flex min-h-14.5 items-center justify-between gap-2.5 bg-[#f6efe3] px-4 py-2 shadow-[0_4px_10px_rgba(0,0,0,0.18)]"
            style={{
              transform: `rotate(${-tile.photoTiltDegrees / 2}deg)`,
              clipPath: homeClipPaths.tileLabels[index],
            }}
          >
            <span className="font-(family-name:--font-just-another-hand) text-[34px] leading-[0.85] tracking-[0.5px] text-[#1f1b18] uppercase">
              {tile.label}
            </span>
            <svg
              width="20"
              height="20"
              viewBox="0 0 24 24"
              fill="none"
              stroke="#1f1b18"
              strokeWidth="2.4"
              strokeLinecap="round"
              strokeLinejoin="round"
              aria-hidden="true"
            >
              <path d="M5 12h14M13 6l6 6-6 6" />
            </svg>
          </span>
        </Link>
      ))}
    </section>
  );
}
