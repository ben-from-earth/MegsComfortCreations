import { washiColors } from '@/components/home/home-content';
import type { ShippedPlace } from '@/components/home/shipped-places';
import {
  placesForRibbon,
  shippedCountryNames,
  shippedStateNames,
} from '@/components/home/shipped-places';
import { getShippedPlaces } from '@/components/home/shipped-places-query';
import './home.css';

const washiPatterns = [
  (color: string) =>
    `repeating-linear-gradient(45deg,${color} 0 8px,color-mix(in srgb,${color} 70%,#fff) 8px 16px)`,
  (color: string) =>
    `radial-gradient(circle,rgba(255,255,255,.55) 2px,transparent 2.5px) 0 0/12px 12px,${color}`,
  (color: string) =>
    `linear-gradient(90deg,rgba(255,255,255,.35) 50%,transparent 50%) 0 0/16px 16px,linear-gradient(rgba(255,255,255,.35) 50%,transparent 50%) 0 0/16px 16px,${color}`,
  (color: string) =>
    `repeating-linear-gradient(90deg,${color} 0 3px,color-mix(in srgb,${color} 80%,#fff) 3px 12px)`,
];

const trackCopies = ['first', 'second'] as const;

function WashiTrack({
  places,
  isReversed = false,
}: {
  places: readonly ShippedPlace[];
  isReversed?: boolean;
}) {
  return (
    <ul
      className={
        isReversed
          ? 'home-ribbon-track home-ribbon-track-reverse'
          : 'home-ribbon-track'
      }
    >
      {trackCopies.map((copy) =>
        places.map((place, index) => {
          const tiltDegrees = (((index * 47) % 101) - 50) / 10;
          const offsetY = ((index * 29) % 21) - 10;
          const color = washiColors[index % washiColors.length];
          return (
            <li
              key={`${copy}-${place.name}`}
              className="home-washi"
              style={{
                background: washiPatterns[index % washiPatterns.length](color),
                transform: `translateY(${offsetY}px) rotate(${tiltDegrees}deg)`,
              }}
            >
              <b className="pt-1 font-(family-name:--font-just-another-hand) text-[38px] leading-none font-normal whitespace-nowrap text-[#1f1b18]">
                {place.name}
              </b>
            </li>
          );
        }),
      )}
    </ul>
  );
}

function Sparkle({ color }: { color: string }) {
  return (
    <svg width="20" height="20" viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M12 0C13 8 16 11 24 12C16 13 13 16 12 24C11 16 8 13 0 12C8 11 11 8 12 0Z"
        fill={color}
      />
    </svg>
  );
}

export default async function ShippedPlacesRibbon() {
  const shippedPlaces = await getShippedPlaces();
  const countries = placesForRibbon(
    shippedCountryNames,
    shippedPlaces.countries,
  );
  const usStates = placesForRibbon(shippedStateNames, shippedPlaces.usStates);

  return (
    <section
      aria-labelledby="shipped-places-heading"
      className="overflow-hidden pt-14 pb-16"
    >
      <div className="flex items-center justify-center gap-5.5 px-10">
        <span aria-hidden="true" className="h-px flex-1 bg-[#b9ab97]" />
        <Sparkle color="#f0a83e" />
        <h2
          id="shipped-places-heading"
          className="m-0 text-center text-[15px] font-extrabold tracking-[3.4px] uppercase"
        >
          Sent with care to{' '}
          <span className="text-[#c4452a]">
            {shippedCountryNames.length} countries
          </span>{' '}
          &amp;{' '}
          <span className="text-[#24493b]">all 50 states</span>
        </h2>
        <Sparkle color="#cf4025" />
        <span aria-hidden="true" className="h-px flex-1 bg-[#b9ab97]" />
      </div>
      <div className="home-ribbon-lanes" aria-hidden="true">
        <WashiTrack places={countries} />
        <WashiTrack places={usStates} isReversed />
      </div>
    </section>
  );
}
