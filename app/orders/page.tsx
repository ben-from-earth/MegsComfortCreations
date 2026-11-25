'use client';

import { useEffect, useState } from 'react';
import { usePathname, useRouter, useSearchParams } from 'next/navigation';
import LoadingWidget from '@/components/shared/loading-widget';
import OrdersFilters, { type OrdersFilterText } from '@/orders/orders-filters';
import OrdersTable from '@/orders/orders-table';
import { MEDIA_TYPES, type MediaType } from 'lib/constants/media-types';
import { trpc } from 'lib/trpc/client';

const TEXT_KEYS = ['name', 'number', 'title', 'min', 'max'] as const;

function isMediaType(value: string | null): value is MediaType {
  return MEDIA_TYPES.some((type) => type === value);
}

function countParam(value: string) {
  if (value.trim() === '') {
    return undefined;
  }
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : undefined;
}

function ordersHref(
  pathname: string,
  searchParams: URLSearchParams,
  updates: Record<string, string | null>,
) {
  const params = new URLSearchParams(searchParams.toString());
  for (const [key, value] of Object.entries(updates)) {
    if (value) {
      params.set(key, value);
    } else {
      params.delete(key);
    }
  }
  const next = params.toString();
  return next ? `${pathname}?${next}` : pathname;
}

function readText(searchParams: URLSearchParams): OrdersFilterText {
  return {
    name: searchParams.get('name') ?? '',
    number: searchParams.get('number') ?? '',
    title: searchParams.get('title') ?? '',
    min: searchParams.get('min') ?? '',
    max: searchParams.get('max') ?? '',
  };
}

export default function OrdersPage() {
  const router = useRouter();
  const pathname = usePathname();
  const searchParams = useSearchParams();
  const textFromUrl = readText(searchParams);
  const urlKey = searchParams.toString();
  const typeParam = searchParams.get('type');
  const type = isMediaType(typeParam) ? typeParam : null;
  const [text, setText] = useState(textFromUrl);

  useEffect(() => {
    setText(readText(new URLSearchParams(urlKey)));
  }, [urlKey]);

  useEffect(() => {
    const current = new URLSearchParams(urlKey);
    const changed = TEXT_KEYS.some(
      (key) => text[key] !== (current.get(key) ?? ''),
    );
    if (!changed) {
      return;
    }
    const timeout = window.setTimeout(() => {
      router.replace(
        ordersHref(
          pathname,
          current,
          Object.fromEntries(TEXT_KEYS.map((key) => [key, text[key] || null])),
        ),
      );
    }, 250);
    return () => window.clearTimeout(timeout);
  }, [text, urlKey, pathname, router]);

  const listQuery = trpc.orders.list.useQuery({
    name: textFromUrl.name || undefined,
    number: textFromUrl.number || undefined,
    title: textFromUrl.title || undefined,
    type: type ?? undefined,
    min: type ? countParam(textFromUrl.min) : undefined,
    max: type ? countParam(textFromUrl.max) : undefined,
  });

  return (
    <main className="mx-auto flex w-full max-w-[1360px] flex-col gap-5 px-10 pt-8 pb-14">
      <div>
        <h1 className="m-0 text-[30px] font-black">Orders</h1>
        {listQuery.data ? (
          <p className="mt-1 mb-0 text-sm font-semibold text-[#6b645d]">
            {listQuery.data.rows.length === listQuery.data.total
              ? `${listQuery.data.total} orders`
              : `${listQuery.data.rows.length} of ${listQuery.data.total} orders`}
          </p>
        ) : null}
      </div>

      <OrdersFilters
        text={text}
        onTextChange={(key, value) =>
          setText((current) => ({ ...current, [key]: value }))
        }
        type={type}
        onTypeChange={(next) =>
          router.replace(
            ordersHref(pathname, searchParams, {
              type: next,
              ...(next === null ? { min: null, max: null } : {}),
            }),
          )
        }
      />

      {listQuery.isLoading ? <LoadingWidget message="Loading orders" /> : null}
      {listQuery.error ? (
        <p className="m-0 text-sm font-semibold text-[#9a3b26]">
          Couldn’t load orders. Try again.
        </p>
      ) : null}
      {listQuery.data ? (
        <OrdersTable
          rows={listQuery.data.rows}
          emptyDatabase={listQuery.data.total === 0}
        />
      ) : null}
    </main>
  );
}
