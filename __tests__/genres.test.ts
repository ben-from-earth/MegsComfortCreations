import { POST as addLinkPOST } from '@//api/genres/addlink/route';
import { GET as GETall, getAllResponse } from '@//api/genres/getall/route';
import { GET as GETforBook } from '@//api/genres/getforbook/route';
import { GET as GETbookswithNoGenre } from '@//api/genres/nogenres/route';
import { GET as GETbookswithGenre } from '@//api/genres/route';
import { POST as unlinkPOST } from '@//api/genres/unlink/route';

import { db } from '@/db/client'; // now your Drizzle db instance
import { books, genresBooks } from '@/db/schema'; // drizzle tables
import { sql } from 'drizzle-orm';

import {
  SuccessfulGenreLinkUnlinkResponse,
  SuccessfulPaginationResponse,
} from 'lib/interfaces/globalInterfaces';
import { NextRequest } from 'next/server';

const serverDomain = process.env.SERVER_BASE_URL;

describe('Connection to database succesful', () => {
  test('can query the database', async () => {
    const result = await db.execute(sql`SELECT 1 + 1 AS result`);
    // drizzle-orm/node-postgres returns { rows }
    expect(result.rows[0].result).toBe(2);
  });
});

const bookData = {
  title: 'Genre Test Title',
  author: 'Book Author',
  pageCount: 100,
  pubYear: 2025,
  spineColor: '#ca2f24ff',
  imageUrls: ['http://testurl.com'],
};

const noGenreBookData = {
  title: 'Genre Test Title 2',
  author: 'Book Author',
  pageCount: 100,
  pubYear: 2025,
  spineColor: '#ca2f24ff',
  imageUrls: ['http://testurl.com'],
};

let bookID: string;

beforeAll(async () => {
  // Insert first book (that will get genres linked)
  const [book] = await db
    .insert(books)
    .values({
      title: bookData.title,
      author: bookData.author,
      pageCount: bookData.pageCount,
      pubYear: bookData.pubYear,
      spineColor: bookData.spineColor,
      imageUrls: bookData.imageUrls,
    })
    .returning();

  bookID = book.id;

  // Insert second book (no genre association)
  await db
    .insert(books)
    .values({
      title: noGenreBookData.title,
      author: noGenreBookData.author,
      pageCount: noGenreBookData.pageCount,
      pubYear: noGenreBookData.pubYear,
      spineColor: noGenreBookData.spineColor,
      imageUrls: noGenreBookData.imageUrls,
    })
    .returning();
});

describe('Link/unlink book to genre', () => {
  test('Link bookID with Science Fiction and Fantasy', async () => {
    const req = new NextRequest(`${serverDomain}/genres/addlink`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ bookID, genres: ['Science Fiction', 'Fantasy'] }),
    });

    const res = await addLinkPOST(req as NextRequest);
    const responseBody: {
      genreResponses: SuccessfulGenreLinkUnlinkResponse[];
    } = await res.json();

    expect(responseBody.genreResponses.length).toEqual(2);
    expect(responseBody.genreResponses[1].genre).toEqual('Fantasy');
  });

  test('Unlink Fantasy from book', async () => {
    const req = new NextRequest(`${serverDomain}/genres/unlink`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ bookID, genres: ['Fantasy'] }),
    });

    const res = await unlinkPOST(req as NextRequest);
    const responseBody: {
      genreResponses: SuccessfulGenreLinkUnlinkResponse[];
    } = await res.json();

    expect(responseBody.genreResponses.length).toEqual(1);
    expect(responseBody.genreResponses[0].message).toEqual(
      'Successful genre unlink',
    );
    expect(responseBody.genreResponses[0].genre).toEqual('Fantasy');
  });
});

describe('Get genres requests', () => {
  test('Get all genres', async () => {
    const res = await GETall();
    const responseBody: getAllResponse = await res.json();

    expect(responseBody.genres).toBeDefined();
    expect(responseBody.genres.length).toBe(20);
  });

  test('Get genres for bookID', async () => {
    const req = new NextRequest(
      `${serverDomain}/genres/getforbook?bookID=${bookID}`,
      { method: 'GET' },
    );

    const res = await GETforBook(req as NextRequest);
    const responseBody: { message: string; genres: string[] } =
      await res.json();

    expect(responseBody.message).toEqual(
      `Successfully grabbed genres for bookID ${bookID}`,
    );
    expect(responseBody.genres.length).toEqual(1);
    expect(responseBody.genres[0]).toEqual('Science Fiction');
  });
});

describe('Get books with/without specific genres', () => {
  test('Get all books with Science Fiction as a genre', async () => {
    const req = new NextRequest(
      `${serverDomain}/genres?genre=Science Fiction&sort=title&limit=3&page=1&ascDesc=asc`,
      { method: 'GET' },
    );

    const res = await GETbookswithGenre(req as NextRequest);
    const responseBody: SuccessfulPaginationResponse = await res.json();

    expect(responseBody.message).toEqual(`Successful database gather`);
    expect(responseBody.total).toEqual(1);
    expect(responseBody.paginatedList[0].title).toEqual('Genre Test Title');
  });

  test('Get books with no genre association', async () => {
    const req = new NextRequest(
      `${serverDomain}/nogenres?sort=title&limit=3&page=1&ascDesc=asc`,
      { method: 'GET' },
    );

    const res = await GETbookswithNoGenre(req as NextRequest);
    const responseBody: SuccessfulPaginationResponse = await res.json();

    expect(responseBody.message).toEqual(`Successful database gather`);
    expect(responseBody.total).toEqual(1);
    expect(responseBody.paginatedList[0].title).toEqual('Genre Test Title 2');
  });
});

afterAll(async () => {
  // Clean up join table first due to FK constraints
  await db.delete(genresBooks);
  await db.delete(books);

  // If you still have a raw pg Pool underneath, close it where you create it,
  // not via Drizzle's `db` instance.
});
