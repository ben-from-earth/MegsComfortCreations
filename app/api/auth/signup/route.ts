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

    const { email, password, firstName, lastName } = body;

    const validation = validate(body, userCreateSchema);
    if (!validation.valid) {
      return handleErrors({ error: 'Validation Error', validation });
    }

    const salt = await bcrypt.genSalt();
    const hashedpass = await bcrypt.hash(password, salt);

    const user = await User.create({
      email,
      password: hashedpass,
      firstName,
      lastName,
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
  } catch (error: unknown) {
    console.error(error);

    // Narrow to something with an optional `code` field
    const e = error as { code?: number | string };

    // Map DB duplicate error to our handler shape
    if (e.code === 23505 || e.code === '23505') {
      return handleErrors({ error: e });
    }

    return handleErrors({ error: e });
  }
}
