import { ApiError } from "@/app/api/api-Errors";
import Genre from "@/lib/database/models/genre";
import {
  GenreLinkUnlinkRequest,
  SuccessfulGenreLinkUnlinkResponse,
} from "@/lib/interfaces/globalInterfaces";
import { NextRequest, NextResponse } from "next/server";

export async function POST(req: NextRequest) {
  const body: GenreLinkUnlinkRequest = await req.json();
  const genreResponses: SuccessfulGenreLinkUnlinkResponse[] = [];
  const bookID = body.bookID;

  for (const genre of body.genres) {
    try {
      await Genre.unlink(genre, bookID);
      genreResponses.push({
        message: "Successful genre unlink",
        genre,
        bookID,
      });
    } catch {
      return new ApiError(
        400,
        "Genre Error",
        "Error connecting to the database and/or genre table"
      ).format();
    }
  }
  return NextResponse.json({ genreResponses }, { status: 201 });
}
