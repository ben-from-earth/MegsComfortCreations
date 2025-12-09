// middleware.ts
import { NextResponse } from 'next/server';
import type { NextRequest } from 'next/server';

// Any routes you want to protect
const PROTECTED_PATHS = ['/profile', '/mediacollector', '/showdatabase'];

export function middleware(req: NextRequest) {
  const { pathname } = req.nextUrl;

  // If the request path is NOT one of the protected ones, let it through
  const isProtected = PROTECTED_PATHS.some((path) => pathname.startsWith(path));

  if (!isProtected) {
    return NextResponse.next();
  }

  // Check for jwt cookie (this is what your login route sets)
  const token = req.cookies.get('jwt')?.value;

  // If no token → redirect to /login
  if (!token) {
    const loginUrl = new URL('/login', req.url);
    // Optional: remember where they were trying to go
    loginUrl.searchParams.set('from', pathname);
    return NextResponse.redirect(loginUrl);
  }

  // If there is a token, just let the request pass through.
  // (If you want to *fully* validate the token, we can upgrade this later
  // using an edge-safe JWT library.)
  return NextResponse.next();
}

// Tell Next.js which routes this middleware applies to
export const config = {
  matcher: [
    '/profile/:path*',
    '/mediacollector/:path*',
    '/showdatabase/:path*',
  ],
};
