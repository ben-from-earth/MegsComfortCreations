import { auth } from 'lib/auth';
import { NextResponse } from 'next/server';

export async function POST(req: Request) {
  try {
    const secret = req.headers.get('x-seed-secret');
    if (!secret || secret !== process.env.SEED_SECRET) {
      return NextResponse.json(
        { ok: false, error: 'Unauthorized' },
        { status: 401 },
      );
    }

    const body = await req.json();
    const { email, password, name } = body ?? {};

    if (!email || !password) {
      return NextResponse.json(
        { ok: false, error: 'Missing email or password' },
        { status: 400 },
      );
    }

    const result = await auth.api.createUser({
      body: {
        email,
        password,
        name,
        role: 'admin',
      },
    });

    // Force a response body no matter what Better Auth returns
    return NextResponse.json({ ok: true, result }, { status: 200 });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ ok: false, error: message }, { status: 500 });
  }
}
