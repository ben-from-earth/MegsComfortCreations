'use client';

import { trpc } from 'lib/trpc/client';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { httpBatchLink, type TRPCClient } from '@trpc/client';
import { ReactNode, useState, createContext, useContext } from 'react';
import superjson from 'superjson';
import type { AppRouter } from 'lib/trpc/routers/_app';

const TRPCRuntimeContext = createContext<{
  client: TRPCClient<AppRouter>;
  queryClient: QueryClient;
} | null>(null);

export function TRPCProvider({ children }: { children: ReactNode }) {
  const [queryClient] = useState(() => new QueryClient());
  const [trpcClient] = useState(() =>
    trpc.createClient({
      links: [httpBatchLink({ url: '/api/trpc', transformer: superjson })],
    }),
  );

  return (
    <trpc.Provider client={trpcClient} queryClient={queryClient}>
      <QueryClientProvider client={queryClient}>
        <TRPCRuntimeContext.Provider
          value={{ client: trpcClient, queryClient }}
        >
          {children}
        </TRPCRuntimeContext.Provider>
      </QueryClientProvider>
    </trpc.Provider>
  );
}

// Convenience hooks to mirror a `createTRPCContext<AppRouter>()` API
// and provide a consistent interface across codebases.
export function useTRPC() {
  return trpc;
}

export function useTRPCClient() {
  const ctx = useContext(TRPCRuntimeContext);
  if (!ctx) throw new Error('useTRPCClient must be used within TRPCProvider');
  return ctx.client;
}
