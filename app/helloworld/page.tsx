'use client';
import { trpc } from '@/lib/trpc/client';

export default function HelloWorld() {
  const { data, isLoading } = trpc.health.hello.useQuery();
  return (
    <div style={{ padding: 16 }}>
      <h1>tRPC Hello World</h1>
      {isLoading ? (
        <p>Loading...</p>
      ) : (
        <pre>{JSON.stringify(data, null, 2)}</pre>
      )}
    </div>
  );
}
