// app/api/auth/signup/route.ts
import { NextRequest, NextResponse } from 'next/server';
import bcrypt from 'bcrypt';
import { validate } from 'jsonschema';

import userCreateSchema from '@/lib/schemas/userCreateSchema.json';
import { User } from '@/lib/database/models/user';
import { createToken, handleErrors, MAX_AGE_SECONDS } from '../_utils';

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();

    const { email, password, first_name, last_name } = body;

    const validation = validate(body, userCreateSchema as any);
    if (!validation.valid) {
      return handleErrors({ error: 'Validation Error', validation });
    }

    const salt = await bcrypt.genSalt();
    const hashedpass = await bcrypt.hash(password, salt);

    const user = await User.create({
      email,
      password: hashedpass,
      first_name,
      last_name,
    });

    if (!user?.id) {
      return handleErrors({ error: 'User Creation Error' });
    }

    const token = createToken(user.id);

    const res = NextResponse.json(
      {
        message: 'Successfully added user',
        user,
      },
      { status: 201 },
    );

    // maxAge in Next.js cookies is in seconds
    res.cookies.set('jwt', token, {
      httpOnly: true,
      maxAge: MAX_AGE_SECONDS,
      path: '/',
    });

    return res;
  } catch (error: any) {
    console.error(error);

    // Map DB duplicate error to our handler shape
    if (error?.code === 23505) {
      return handleErrors({ error });
    }

    return handleErrors({ error });
  }
}
