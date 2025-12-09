// app/api/auth/me/route.ts
import { NextRequest, NextResponse } from 'next/server';

import { verifyTokenWithId } from '../_utils';
import { User, UserType } from '@/lib/database/models/user';

export type SuccessfulUserResponse = {
  access: true;
  user: UserType;
};

export async function GET(req: NextRequest) {
  const token = req.cookies.get('jwt')?.value;

  if (!token) {
    return NextResponse.json(
      {
        error: 'Access Denied',
        message: 'This route is protected, please sign in or register',
      },
      { status: 403 },
    );
  }

  try {
    const decoded = verifyTokenWithId(token);
    const user = await User.findOne(decoded.id);

    return NextResponse.json(
      {
        access: true,
        user,
      },
      { status: 200 },
    );
  } catch (error) {
    console.error(error);
    return NextResponse.json(
      {
        error: 'Access Denied',
        message: 'This route is protected, please sign in or register',
      },
      { status: 403 },
    );
  }
}
