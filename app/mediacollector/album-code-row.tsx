'use client';

import { useTransition } from 'react';
import Image from 'next/image';
import { trpc } from 'lib/trpc/client';
import {
  albumCodePair,
  albumCodeStatus,
  isSameAlbumCodePair,
  type AlbumCodePair,
  type AlbumCodeResult,
  type AlbumCodeStatus,
} from 'lib/helpers/album-code-status';
import type { SpotifyAlbumResult } from 'lib/trpc/routers/collect/actions/get-spotify-album';
import { spotifyCodeUrl } from 'lib/trpc/routers/collect/actions/spotify-code';
import type { MediaItemForm } from './collector-form/media-item-form-schema';

export type AlbumCodeLookup =
  | {
      found: true;
      uri: string;
      coverUrl: string | null;
      pair: AlbumCodePair;
    }
  | { found: false; pair: AlbumCodePair };

export function applyAlbumCodeLookup(
  item: MediaItemForm,
  lookup: AlbumCodeLookup,
): MediaItemForm {
  if (!lookup.found) {
    return {
      ...item,
      blockInfo: {
        ...item.blockInfo,
        spotifyAlbumUri: null,
        codeStatus: 'noMatch',
        codeGeneratedFor: lookup.pair,
      },
    };
  }
  return {
    ...item,
    images:
      item.images.length === 0 && lookup.coverUrl
        ? [
            {
              url: lookup.coverUrl,
              selected: true,
              isDefault: true,
              spineColor: '#ffffff',
              codeColor: null,
            },
          ]
        : item.images,
    blockInfo: {
      ...item.blockInfo,
      spotifyAlbumUri: lookup.uri,
      codeStatus: 'generated',
      codeGeneratedFor: lookup.pair,
    },
  };
}

const NOT_FOUND: SpotifyAlbumResult = { found: false };

const pillClassName =
  'inline-flex h-7 shrink-0 items-center justify-center gap-1.5 rounded-full border text-xs font-extrabold disabled:cursor-not-allowed';

function spotifyAlbumPageUrl(uri: string) {
  const id = uri.split(':').at(-1);
  return id ? `https://open.spotify.com/album/${id}` : null;
}

function statusText(status: AlbumCodeStatus, title: string, artist: string) {
  switch (status) {
    case 'missingName':
      return 'Add a title and artist first';
    case 'missingColor':
      return 'Pick a code color for this cover first';
    case 'ready':
      return 'Ready to generate';
    case 'generating':
      return 'Searching Spotify…';
    case 'generated':
      return `${title} · ${artist}`;
    case 'generatedWithoutColor':
      return 'Pick a code color for this cover';
    case 'outOfDate':
      return 'Out of date · title or artist changed';
    case 'noMatch':
      return `No Spotify match for ${title}`;
  }
}

function GenerateIcon() {
  return (
    <svg
      width="14"
      height="14"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.2"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M5 19L15 9" />
      <path d="M15 9l2-2" />
      <path d="M18 2v3M16.5 3.5h3M20 9v2M19 10h2M11 3v2M10 4h2" />
    </svg>
  );
}

function RegenerateIcon() {
  return (
    <svg
      width="14"
      height="14"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.2"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M20 12a8 8 0 11-2.3-5.6" />
      <path d="M20 4v5h-5" />
    </svg>
  );
}

export interface AlbumCodeRowProps {
  title: string;
  artist: string | null | undefined;
  releaseDate: string | null | undefined;
  spotifyAlbumUri: string | null | undefined;
  codeStatus: AlbumCodeResult | undefined;
  codeGeneratedFor: AlbumCodePair | null | undefined;
  codeColor: string | null | undefined;
  isLocked: boolean;
  onLookup: (lookup: AlbumCodeLookup) => void;
}

export default function AlbumCodeRow({
  title,
  artist,
  releaseDate,
  spotifyAlbumUri,
  codeStatus,
  codeGeneratedFor,
  codeColor,
  isLocked,
  onLookup,
}: AlbumCodeRowProps) {
  const [isGenerating, startGenerating] = useTransition();
  const { mutateAsync: findSpotifyAlbum } =
    trpc.collect.findSpotifyAlbum.useMutation();
  const pair = albumCodePair(title, artist);
  const status = albumCodeStatus({
    title,
    artist,
    codeColor,
    spotifyAlbumUri,
    codeStatus,
    codeGeneratedFor,
    isGenerating,
  });
  const previewUrl =
    spotifyAlbumUri && codeColor
      ? spotifyCodeUrl(spotifyAlbumUri, codeColor)
      : null;
  const albumPageUrl = spotifyAlbumUri
    ? spotifyAlbumPageUrl(spotifyAlbumUri)
    : null;
  const showsPreview =
    previewUrl !== null && (status === 'generated' || status === 'outOfDate');
  const canGenerate =
    !isGenerating &&
    pair.title !== '' &&
    pair.artist !== '' &&
    Boolean(codeColor);
  const text =
    !isLocked || status === 'generated'
      ? statusText(status, pair.title, pair.artist)
      : spotifyAlbumUri
        ? 'No code color on this cover'
        : 'No code yet';
  const disabledReason =
    pair.title === '' || pair.artist === ''
      ? 'Add a title and artist first'
      : 'Pick a code color for this cover first';

  const runLookup = (shouldSearch: boolean) => {
    if (!shouldSearch && spotifyAlbumUri) {
      onLookup({ found: true, uri: spotifyAlbumUri, coverUrl: null, pair });
      return;
    }
    startGenerating(async () => {
      const result = await findSpotifyAlbum({
        title: pair.title,
        artist: pair.artist,
        releaseDate: releaseDate ?? null,
      }).catch(() => NOT_FOUND);
      onLookup(
        result.found
          ? { found: true, uri: result.uri, coverUrl: result.coverUrl, pair }
          : { found: false, pair },
      );
    });
  };
  const handleGenerate = () =>
    runLookup(!spotifyAlbumUri || !isSameAlbumCodePair(codeGeneratedFor, pair));

  return (
    <div
      role="group"
      aria-label="Spotify code"
      aria-live="polite"
      className={`mx-4.5 mb-3.5 flex h-10 flex-nowrap items-center gap-2 rounded-[10px] pr-1.5 pl-2.5 ${
        status === 'outOfDate'
          ? 'bg-[#fffaf0] shadow-[inset_0_0_0_1.5px_#c9973f]'
          : status === 'noMatch'
            ? 'bg-[#f9eee9] shadow-[inset_0_0_0_1.5px_#d9a493]'
            : 'bg-[#f7f4f0]'
      }`}
    >
      <span className="shrink-0 text-[11px] font-extrabold tracking-[0.5px] text-[#6b645d] uppercase">
        Code
      </span>
      {status === 'generating' ? (
        <div
          aria-hidden="true"
          className="h-4 w-4 shrink-0 animate-spin rounded-full border-2 border-[#d6cec4] border-t-[#2b2826] motion-reduce:animate-none"
        />
      ) : showsPreview ? (
        <a
          href={albumPageUrl ?? undefined}
          target="_blank"
          rel="noreferrer"
          title={`Open ${pair.title} on Spotify`}
          className={`relative block h-7 w-30 shrink-0 overflow-hidden rounded-[5px] shadow-[inset_0_0_0_1px_rgba(0,0,0,0.12)] ${
            status === 'outOfDate' ? 'opacity-50' : ''
          }`}
          style={{ backgroundColor: codeColor ?? undefined }}
        >
          <Image
            src={previewUrl}
            alt="Spotify code preview"
            fill
            sizes="120px"
            unoptimized
            loader={({ src }) => src}
            className="object-contain"
          />
        </a>
      ) : null}
      <span
        title={
          status === 'generated'
            ? `Generated from ${pair.title} by ${pair.artist}`
            : undefined
        }
        className={`min-w-0 flex-1 truncate text-xs font-bold ${
          status === 'outOfDate'
            ? 'text-[#8a5a22]'
            : status === 'noMatch'
              ? 'text-[#9a3b26]'
              : 'text-[#6b645d]'
        }`}
      >
        {text}
      </span>
      {isLocked ? null : status === 'generated' ||
        status === 'generatedWithoutColor' ? (
        <button
          type="button"
          aria-label="Regenerate code"
          title={canGenerate ? 'Regenerate the code' : disabledReason}
          disabled={!canGenerate}
          onClick={() => runLookup(true)}
          className={`${pillClassName} w-7 cursor-pointer border-[#e0d9d0] bg-white p-0 text-[#2b2826] disabled:text-[#a59d93]`}
        >
          <RegenerateIcon />
        </button>
      ) : status === 'outOfDate' || status === 'noMatch' ? (
        <button
          type="button"
          title={canGenerate ? undefined : disabledReason}
          disabled={!canGenerate}
          onClick={handleGenerate}
          className={`${pillClassName} cursor-pointer border-[#c9973f] bg-white px-3 text-[#6b4a14] disabled:border-[#ddd5cb] disabled:text-[#a59d93]`}
        >
          <RegenerateIcon />
          {status === 'outOfDate' ? 'Regenerate' : 'Try again'}
        </button>
      ) : (
        <button
          type="button"
          aria-label="Generate code"
          title={canGenerate ? 'Generate the code' : disabledReason}
          disabled={!canGenerate}
          onClick={handleGenerate}
          className={`${pillClassName} cursor-pointer border-transparent bg-[#2b2826] px-3 text-white disabled:border-[#ddd5cb] disabled:bg-[#f4f1ec] disabled:text-[#a59d93]`}
        >
          <GenerateIcon />
          Generate
        </button>
      )}
    </div>
  );
}
