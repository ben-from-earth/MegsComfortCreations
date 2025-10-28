import { NextResponse } from "next/server";

export async function GET() {
  const user = {
    id: 123,
    firstName: "Ben",
    lastName: "Knox",
    email: "example@email.com",
  };

  return NextResponse.json(user);
}
