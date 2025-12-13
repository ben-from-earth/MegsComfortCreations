import { NextRequest, NextResponse } from 'next/server';
import bcrypt from 'bcrypt';
import { User } from '@/lib/database/models/user';

import { createToken, handleErrors, MAX_AGE_SECONDS } from '../_utils';

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    console.log('Login body:', body);
    const { email, password } = body;

    const user = await User.login(email);

    if (!user) {
      return handleErrors({ error: 'Email Incorrect', email });
    }

    const auth = await bcrypt.compare(password, user.password_hash);

    if (!auth) {
      return handleErrors({ error: 'Incorrect Password' });
    }

    const token = createToken(user.id);

    const res = NextResponse.json(
      { user, message: 'Successful login' },
      { status: 200 },
    );

    res.cookies.set('jwt', token, {
      httpOnly: true,
      maxAge: MAX_AGE_SECONDS,
      path: '/',
    });

    return res;
  } catch (error: unknown) {
    console.error(error);
    return handleErrors({ error });
  }
}
