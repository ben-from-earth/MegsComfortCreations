import { NextResponse } from 'next/server';

export async function GET() {
  const res = NextResponse.json(
    { access: false, message: 'Successful logout' },
    { status: 200 },
  );

  // Clear JWT cookie
  res.cookies.set('jwt', '', {
    httpOnly: true,
    maxAge: 0,
    path: '/',
  });

  return res;
}
