// app/api/auth/_utils.ts
import jwt, { Secret } from 'jsonwebtoken';
import { NextResponse } from 'next/server';

const envSecret = process.env.JWTSECRET;

if (!envSecret) {
  throw new Error('JWTSECRET environment variable is not set');
}

// ✅ Now JWT_SECRET is guaranteed to be a valid Secret at both runtime and type level
export const JWT_SECRET: Secret = envSecret;

// 3 days in seconds
export const MAX_AGE_SECONDS = 3 * 24 * 60 * 60;

export const createToken = (id: string | number) => {
  return jwt.sign({ id }, JWT_SECRET, {
    expiresIn: MAX_AGE_SECONDS,
  });
};

type ValidationError = {
  error: 'Validation Error';
  validation: {
    errors: { stack: string }[];
  };
};

type IncorrectPasswordError = {
  error: 'Incorrect Password';
};

type EmailIncorrectError = {
  error: 'Email Incorrect';
  email: string;
};

type UserCreationError = {
  error: 'User Creation Error' | 'User Creation error';
};

type DuplicateError = {
  error: {
    code: number;
  };
};

type GenericError = {
  error: any;
  email?: string;
};

type AppError =
  | ValidationError
  | IncorrectPasswordError
  | EmailIncorrectError
  | UserCreationError
  | DuplicateError
  | GenericError;

export function handleErrors(err: AppError): NextResponse {
  if ((err as DuplicateError).error?.code === 23505) {
    return NextResponse.json(
      {
        error: 'Duplication Error',
        message: 'That email is already registered',
      },
      { status: 400 },
    );
  }

  if (
    err.error === 'User Creation error' ||
    err.error === 'User Creation Error'
  ) {
    return NextResponse.json(
      {
        error: 'User Creation Error',
        message:
          'There was a user creation error or the database is currently offline',
      },
      { status: 400 },
    );
  }

  if (err.error === 'Validation Error') {
    const validationErr = err as ValidationError;
    return NextResponse.json(
      {
        error: 'Validation Error',
        errors: validationErr.validation.errors.map((e) => e.stack),
      },
      { status: 400 },
    );
  }

  if (err.error === 'Incorrect Password') {
    return NextResponse.json(
      {
        error: 'Authentication Error',
        message: 'Inputted password is incorrect',
      },
      { status: 403 },
    );
  }

  if (err.error === 'Email Incorrect') {
    const emailErr = err as EmailIncorrectError;
    return NextResponse.json(
      {
        error: 'Authentication Error',
        message: `Inputted email (${emailErr.email}) is not in the userbase`,
      },
      { status: 403 },
    );
  }

  console.error('Unhandled error in handleErrors:', err);

  return NextResponse.json(
    {
      error: 'Server Error',
      message: 'An unexpected error occurred',
    },
    { status: 500 },
  );
}

export type JwtPayloadWithId = {
  id: string;
} & jwt.JwtPayload;

export function verifyTokenWithId(token: string): JwtPayloadWithId {
  const decoded = jwt.verify(token, JWT_SECRET);

  // Runtime check to make TS happy and keep it safe
  if (!decoded || typeof decoded !== 'object') {
    throw new Error('Invalid token payload');
  }

  if (!('id' in decoded)) {
    throw new Error('Token payload missing id');
  }

  return decoded as JwtPayloadWithId;
}
