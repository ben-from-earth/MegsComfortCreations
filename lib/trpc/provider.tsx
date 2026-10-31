'use client';

import { trpc } from 'lib/trpc/client';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { httpBatchLink, httpBatchStreamLink, splitLink } from '@trpc/client';
import { ReactNode, useState } from 'react';
import superjson from 'superjson';

export function TRPCProvider({ children }: { children: ReactNode }) {
  const [queryClient] = useState(() => new QueryClient());
  const [trpcClient] = useState(() =>
    trpc.createClient({
      links: [
        splitLink({
          condition: (operation) =>
            operation.path === 'collect.collectMediaStream',
          true: httpBatchStreamLink({
            url: '/api/trpc',
            transformer: superjson,
          }),
          false: httpBatchLink({ url: '/api/trpc', transformer: superjson }),
        }),
      ],
    }),
  );

  return (
    <trpc.Provider client={trpcClient} queryClient={queryClient}>
      <QueryClientProvider client={queryClient}>{children}</QueryClientProvider>
    </trpc.Provider>
  );
}
