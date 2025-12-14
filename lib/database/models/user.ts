import { users } from '@/app/db/schema';
import { db } from '@/app/db/client';
import { eq } from 'drizzle-orm';

export class User {
  static async create({
    email,
    password,
    firstName,
    lastName,
  }: {
    email: string;
    password: string;
    firstName: string;
    lastName: string;
  }): Promise<UserType> {
    const [user] = await db
      .insert(users)
      .values({
        email,
        passwordHash: password,
        firstName,
        lastName,
      })
      .returning({
        id: users.id,
        email: users.email,
        firstName: users.firstName,
        lastName: users.lastName,
      });

    return user;
  }

  static async login(email: string): Promise<UserLoginRow | undefined> {
    const [user] = await db
      .select({
        id: users.id,
        email: users.email,
        firstName: users.firstName,
        lastName: users.lastName,
        passwordHash: users.passwordHash,
      })
      .from(users)
      .where(eq(users.email, email));

    return user;
  }

  static async findOne(id: string): Promise<UserType | undefined> {
    const [user] = await db
      .select({
        id: users.id,
        email: users.email,
        firstName: users.firstName,
        lastName: users.lastName,
      })
      .from(users)
      .where(eq(users.id, id));

    return user;
  }
}

export type UserType = {
  id: string;
  email: string;
  firstName: string;
  lastName: string;
};

export type UserLoginRow = UserType & {
  passwordHash: string;
};
