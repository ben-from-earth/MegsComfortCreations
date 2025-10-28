import { POST as addLinkPOST } from "@/app/api/genres/addlink/route";
import { GET as GETall, getAllResponse } from "@/app/api/genres/getall/route";
import { GET as GETforBook } from "@/app/api/genres/getforbook/route";
import { GET as GETbookswithNoGenre } from "@/app/api/genres/nogenres/route";
import { GET as GETbookswithGenre } from "@/app/api/genres/route";
import { POST as unlinkPOST } from "@/app/api/genres/unlink/route";
import db from "@/lib/database/db";
import {
  postSavedMediaItem,
  SuccessfulGenreLinkUnlinkResponse,
  SuccessfulPaginationResponse,
} from "@/lib/interfaces/globalInterfaces";
import { NextRequest } from "next/server";

const serverDomain = process.env.SERVER_BASE_URL;

describe("Connection to database succesful", () => {
  test("can query the database", async () => {
    const result = await db.query<{ result: number }>("SELECT 1 + 1 AS result");
    expect(result.rows[0].result).toBe(2);
  });
});

const bookData = {
  title: "Genre Test Title",
  author: "Book Author",
  page_count: 100,
  pub_year: 2025,
  spine_color: "#ca2f24ff",
  image_urls: ["http://testurl.com"],
};
const noGenreBookData = {
  title: "Genre Test Title 2",
  author: "Book Author",
  page_count: 100,
  pub_year: 2025,
  spine_color: "#ca2f24ff",
  image_urls: ["http://testurl.com"],
};
let bookID: string;

beforeAll(async function () {
  const bookResult = await db.query<postSavedMediaItem>(
    `INSERT INTO books (
            title,
            author,
            page_count,
            pub_year,
            spine_color,
            image_urls             
      ) VALUES ($1, $2, $3, $4, $5, $6) 
      RETURNING 
            id,
            title,
            author,
            page_count,
            pub_year,
            spine_color,
            image_urls`,
    [
      bookData.title,
      bookData.author,
      bookData.page_count,
      bookData.pub_year,
      bookData.spine_color,
      bookData.image_urls,
    ]
  );

  bookID = bookResult.rows[0].id;
  await db.query<postSavedMediaItem>(
    `INSERT INTO books (
            title,
            author,
            page_count,
            pub_year,
            spine_color,
            image_urls             
      ) VALUES ($1, $2, $3, $4, $5, $6) 
      RETURNING 
            id,
            title,
            author,
            page_count,
            pub_year,
            spine_color,
            image_urls`,
    [
      noGenreBookData.title,
      noGenreBookData.author,
      noGenreBookData.page_count,
      noGenreBookData.pub_year,
      noGenreBookData.spine_color,
      noGenreBookData.image_urls,
    ]
  );
});

describe("Link/unlink book to genre", () => {
  test("Link bookID with Science Fiction and Fantasy", async () => {
    const req = new NextRequest(`${serverDomain}/genres/addlink`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ bookID, genres: ["Science Fiction", "Fantasy"] }),
    });

    const res = await addLinkPOST(req as NextRequest);
    const responseBody: {
      genreResponses: SuccessfulGenreLinkUnlinkResponse[];
    } = await res.json();

    expect(responseBody.genreResponses.length).toEqual(2);
    expect(responseBody.genreResponses[1].genre).toEqual("Fantasy");
  });

  test("Unlink Fantasy from book", async () => {
    const req = new NextRequest(`${serverDomain}/genres/unlink`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ bookID, genres: ["Fantasy"] }),
    });

    const res = await unlinkPOST(req as NextRequest);
    const responseBody: {
      genreResponses: SuccessfulGenreLinkUnlinkResponse[];
    } = await res.json();
    expect(responseBody.genreResponses.length).toEqual(1);
    expect(responseBody.genreResponses[0].message).toEqual(
      "Successful genre unlink"
    );
    expect(responseBody.genreResponses[0].genre).toEqual("Fantasy");
  });
});

describe("Get genres requests", () => {
  test("Get all genres", async () => {
    const res = await GETall();
    const responseBody: getAllResponse = await res.json();

    expect(responseBody.genres).toBeDefined();
    expect(responseBody.genres.length).toBe(20);
  });

  test("Get genres for bookID", async () => {
    const req = new NextRequest(
      `${serverDomain}/genres/getforbook?bookID=${bookID}`,
      { method: "GET" }
    );

    const res = await GETforBook(req as NextRequest);
    const responseBody: { message: string; genres: string[] } =
      await res.json();

    expect(responseBody.message).toEqual(
      `Successfully grabbed genres for bookID ${bookID}`
    );
    expect(responseBody.genres.length).toEqual(1);
    expect(responseBody.genres[0]).toEqual("Science Fiction");
  });
});

describe("Get books with/without specific genres", () => {
  test("Get all books with Science Fiction as a genre", async () => {
    const req = new NextRequest(
      `${serverDomain}/genres?genre=Science Fiction&sort=title&limit=3&page=1&ascDesc=asc`,
      { method: "GET" }
    );

    const res = await GETbookswithGenre(req as NextRequest);
    const responseBody: SuccessfulPaginationResponse = await res.json();

    expect(responseBody.message).toEqual(`Successful database gather`);
    expect(responseBody.total).toEqual(1);
    expect(responseBody.paginatedList[0].title).toEqual("Genre Test Title");
  });

  test("Get books with no genre association", async () => {
    const req = new NextRequest(
      `${serverDomain}/nogenres?sort=title&limit=3&page=1&ascDesc=asc`,
      { method: "GET" }
    );

    const res = await GETbookswithNoGenre(req as NextRequest);
    const responseBody: SuccessfulPaginationResponse = await res.json();

    expect(responseBody.message).toEqual(`Successful database gather`);
    expect(responseBody.total).toEqual(1);
    expect(responseBody.paginatedList[0].title).toEqual("Genre Test Title 2");
  });
});

afterAll(async function () {
  await db.query("DELETE FROM genres_books");
  await db.query("DELETE FROM books");
  await db.end();
});
