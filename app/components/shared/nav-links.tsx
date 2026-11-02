'use client';

import Image from 'next/image';
import Link from 'next/link';
import { Nunito } from 'next/font/google';
import { usePathname } from 'next/navigation';
import { useEffect, useId, useLayoutEffect, useRef, useState } from 'react';
import { flushSync, preload } from 'react-dom';

const nunito = Nunito({ weight: '800', subsets: ['latin'] });

type Brush = { src: string; ink: string; tilt: number };

const BRUSHES = [
  { src: '/brushes/brush-2.png', ink: '#2b2826', tilt: 1 },
  { src: '/brushes/brush-3.png', ink: '#2b2826', tilt: -1 },
  { src: '/brushes/brush-4.png', ink: '#ffffff', tilt: 0.5 },
  { src: '/brushes/brush-5.png', ink: '#2b2826', tilt: 1.5 },
  { src: '/brushes/brush-6.png', ink: '#2b2826', tilt: -0.5 },
  { src: '/brushes/brush-7.png', ink: '#2b2826', tilt: 1 },
] satisfies readonly Brush[];

const SHOP_LINKS = [
  { href: '/shop/custom-ornaments', label: 'Custom Ornaments' },
  { href: '/shop/stickers', label: 'Stickers' },
  { href: '/shop/bookmarks', label: 'Bookmarks' },
  { href: '/shop/miniatures', label: 'Miniatures' },
  { href: '/shop/3d-printing', label: '3D Printing' },
] as const;

const PUBLIC_LINKS = [
  { href: '/blog', label: 'Blog' },
  { href: '/about', label: 'About' },
] as const;

const ADMIN_LINKS = [
  { href: '/mediacollector', label: 'Collector' },
  { href: '/library', label: 'Library' },
] as const;

const TAB_CLASS =
  'relative inline-flex h-11 cursor-pointer items-center gap-1.5 border-0 bg-transparent px-[18px] text-[13px] leading-none tracking-[1.2px] text-[#2b2826] uppercase no-underline focus-visible:rounded-lg focus-visible:outline-3 focus-visible:outline-offset-2 focus-visible:outline-[#3d3935]';

function isCurrentRoute(pathname: string, href: string) {
  return pathname === href || pathname.startsWith(`${href}/`);
}

function TabLabel({ label, brush }: { label: string; brush: Brush | null }) {
  return (
    <>
      {brush ? (
        <Image
          src={brush.src}
          alt=""
          aria-hidden="true"
          width={640}
          height={220}
          unoptimized
          loading="eager"
          className="pointer-events-none absolute -top-1.5 -left-3.5 h-[calc(100%+12px)] w-[calc(100%+28px)] max-w-none object-fill"
          style={{ transform: `rotate(${brush.tilt}deg)` }}
        />
      ) : null}
      <span
        className="relative"
        style={brush ? { color: brush.ink } : undefined}
      >
        {label}
      </span>
    </>
  );
}

export default function NavLinks({ isAdmin }: { isAdmin: boolean }) {
  const pathname = usePathname();
  const shopMenuId = useId();
  const [brushIndex, setBrushIndex] = useState<number | null>(null);
  const [isShopOpen, setIsShopOpen] = useState(false);
  const shopWrapperRef = useRef<HTMLSpanElement>(null);
  const shopButtonRef = useRef<HTMLButtonElement>(null);
  const shopMenuRef = useRef<HTMLUListElement>(null);

  for (const { src } of BRUSHES) preload(src, { as: 'image' });

  useLayoutEffect(() => {
    setBrushIndex((previous) => {
      if (previous === null) return Math.floor(Math.random() * BRUSHES.length);
      const next = Math.floor(Math.random() * (BRUSHES.length - 1));
      return next >= previous ? next + 1 : next;
    });
  }, [pathname]);

  useEffect(() => {
    if (!isShopOpen) return;
    const closeOnOutsideClick = (event: PointerEvent) => {
      if (
        event.target instanceof Node &&
        !shopWrapperRef.current?.contains(event.target)
      ) {
        setIsShopOpen(false);
      }
    };
    document.addEventListener('pointerdown', closeOnOutsideClick);
    return () =>
      document.removeEventListener('pointerdown', closeOnOutsideClick);
  }, [isShopOpen]);

  const brush = brushIndex === null ? null : BRUSHES[brushIndex];
  const isShopCurrent = SHOP_LINKS.some((link) =>
    isCurrentRoute(pathname, link.href),
  );

  const getShopItems = () =>
    Array.from(
      shopMenuRef.current?.querySelectorAll<HTMLAnchorElement>(
        '[role="menuitem"]',
      ) ?? [],
    );

  function handleShopButtonKeyDown(
    event: React.KeyboardEvent<HTMLButtonElement>,
  ) {
    if (event.key !== 'ArrowDown') return;
    event.preventDefault();
    flushSync(() => setIsShopOpen(true));
    getShopItems()[0]?.focus();
  }

  function handleShopKeyDown(event: React.KeyboardEvent<HTMLSpanElement>) {
    if (event.key === 'Escape' && isShopOpen) {
      setIsShopOpen(false);
      shopButtonRef.current?.focus();
      return;
    }
    if (!shopMenuRef.current?.contains(document.activeElement)) return;

    const items = getShopItems();
    const focusedIndex = items.findIndex(
      (item) => item === document.activeElement,
    );
    let targetIndex: number;
    switch (event.key) {
      case 'ArrowDown':
        targetIndex = (focusedIndex + 1) % items.length;
        break;
      case 'ArrowUp':
        targetIndex = (focusedIndex - 1 + items.length) % items.length;
        break;
      case 'Home':
        targetIndex = 0;
        break;
      case 'End':
        targetIndex = items.length - 1;
        break;
      default:
        return;
    }
    event.preventDefault();
    items[targetIndex]?.focus();
  }

  return (
    <nav
      aria-label="Main"
      className={`${nunito.className} flex items-center gap-1.5`}
    >
      <span
        ref={shopWrapperRef}
        className="relative inline-flex"
        onKeyDown={handleShopKeyDown}
      >
        <button
          ref={shopButtonRef}
          type="button"
          className={TAB_CLASS}
          aria-haspopup="true"
          aria-expanded={isShopOpen}
          aria-controls={shopMenuId}
          onClick={() => setIsShopOpen((open) => !open)}
          onKeyDown={handleShopButtonKeyDown}
        >
          <TabLabel label="Shop" brush={isShopCurrent ? brush : null} />
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2.6"
            strokeLinecap="round"
            strokeLinejoin="round"
            aria-hidden="true"
            className={`relative transition-transform duration-150 ${isShopOpen ? 'rotate-180' : ''}`}
            style={isShopCurrent && brush ? { color: brush.ink } : undefined}
          >
            <path d="M6 9l6 6 6-6" />
          </svg>
        </button>
        <ul
          ref={shopMenuRef}
          id={shopMenuId}
          role="menu"
          aria-label="Shop"
          hidden={!isShopOpen}
          className="absolute top-[calc(100%+10px)] left-0 z-20 m-0 min-w-[230px] list-none rounded-[14px] border border-[#e3e9de] bg-white p-2 shadow-[0_14px_32px_rgba(40,45,35,.16)]"
        >
          {SHOP_LINKS.map((link) => (
            <li key={link.href} role="none">
              <Link
                href={link.href}
                role="menuitem"
                aria-current={
                  isCurrentRoute(pathname, link.href) ? 'page' : undefined
                }
                onClick={() => setIsShopOpen(false)}
                className="flex h-10 items-center rounded-lg px-3.5 text-[13px] tracking-[1px] text-[#2b2826] uppercase no-underline outline-none hover:bg-[#eef2ea] focus-visible:bg-[#eef2ea]"
              >
                {link.label}
              </Link>
            </li>
          ))}
        </ul>
      </span>

      {(isAdmin ? [...PUBLIC_LINKS, ...ADMIN_LINKS] : PUBLIC_LINKS).map(
        (link) => {
          const isCurrent = isCurrentRoute(pathname, link.href);
          return (
            <Link
              key={link.href}
              href={link.href}
              aria-current={isCurrent ? 'page' : undefined}
              className={TAB_CLASS}
            >
              <TabLabel label={link.label} brush={isCurrent ? brush : null} />
            </Link>
          );
        },
      )}
    </nav>
  );
}
