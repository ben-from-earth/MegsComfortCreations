import { db } from '@/db/client';
import {
  MediaType,
  PostSavedMediaItem,
  PreSavedMediaItem,
  SuccessfulMediaSaveEditResponse,
  SuccessfulMediaSearchResponse,
  SuccessfulPaginationResponse,
} from 'lib/interfaces/globalInterfaces';

import { GET as databaseSearchGET } from '@//api/database/search/route';
import { GET as paginationGET } from '@//api/database/route';
import { DELETE } from '@//api/database/delete/route';
import { POST } from '@//api/database/save/[type]/route';
import { PUT } from '@//api/database/edit/[type]/route';

import { NextRequest } from 'next/server';
import {
  ApiError,
  DatabaseSaveEditErrorResponse,
  ErrorResponse,
  SearchErrorResponse,
} from '../app/api/api-Errors';
import { randomUUID } from 'crypto';

// drizzle
import { sql } from 'drizzle-orm';
import { books, movies, videoGames, albums } from 'lib/database/schema';

const serverDomain = process.env.SERVER_BASE_URL;

describe('Connection to database succesful', () => {
  test('can query the database', async () => {
    const result = await db.execute(sql`SELECT 1 + 1 AS result`);
    // For node-postgres driver, execute() returns QueryResult with .rows
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((result as any).rows[0].result).toBe(2);
  });
});

const otherMedias: MediaType[] = ['movie', 'video_game', 'album'];
const saveIDs: Record<string, string> = {
  bookID: '',
  movieID: '',
  video_gameID: '',
  albumID: '',
};

const testBookReq = {
  title: 'Dune',
  author: 'Frank Herbert',
  pageCount: 584,
  pubYear: 1965,
  spineColor: '#f25b26',
  imageUrls: ['http://testurl.com'],
};

const testOtherMediaReq = {
  title: 'Media Title',
  spineColor: '#f25b26',
  imageUrls: ['http://testurl.com'],
};

const bookData: PreSavedMediaItem & { id?: string } = {
  title: 'Book Title',
  author: 'Book Author',
  pageCount: 100,
  pubYear: 2025,
  spineColor: '#ca2f24ff',
  imageUrls: ['http://testurl.com'],
};

const otherData: PreSavedMediaItem & { id?: string } = {
  title: 'Other Title',
  spineColor: '#ca2f24ff',
  imageUrls: ['http://testurl.com'],
};

beforeAll(async () => {
  // Insert book
  const [bookResult] = await db
    .insert(books)
    .values({
      title: bookData.title,
      author: bookData.author!,
      pageCount: bookData.pageCount!,
      pubYear: bookData.pubYear!,
      spineColor: bookData.spineColor!,
      imageUrls: bookData.imageUrls,
    })
    .returning();

  saveIDs.bookID = bookResult.id;

  // Insert movie
  const [movieResult] = await db
    .insert(movies)
    .values({
      title: otherData.title,
      spineColor: otherData.spineColor!,
      imageUrls: otherData.imageUrls,
    })
    .returning();

  saveIDs.movieID = movieResult.id;

  // Insert video game
  const [VGResult] = await db
    .insert(videoGames)
    .values({
      title: otherData.title,
      spineColor: otherData.spineColor!,
      imageUrls: otherData.imageUrls,
    })
    .returning();

  saveIDs.video_gameID = VGResult.id;

  // Insert album
  const [albumResult] = await db
    .insert(albums)
    .values({
      title: otherData.title,
      spineColor: otherData.spineColor!,
      imageUrls: otherData.imageUrls,
    })
    .returning();

  saveIDs.albumID = albumResult.id;
});

describe('Test saving book to database', () => {
  test('Attempt database save with proper inputs', async () => {
    const req = new NextRequest(`${serverDomain}/database/save`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(testBookReq),
    });

    const res = await POST(req as NextRequest, { params: { type: 'book' } });
    const responseBody: SuccessfulMediaSaveEditResponse = await res.json();

    const newBook = responseBody.actionAttemptItem;
    const message = responseBody.message;
    expect(newBook.id).toBeDefined();
    expect(message).toEqual('Dune successfully added to database.');
    expect(res.status).toEqual(201);
  });

  test('Attempt save of existing title/author combo', async () => {
    const req = new NextRequest(`${serverDomain}/database/save`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(testBookReq),
    });

    const res = await POST(req as NextRequest, { params: { type: 'book' } });
    const responseBody: DatabaseSaveEditErrorResponse = await res.json();

    expect(responseBody.error).toEqual('Duplication Attempt Error');
    expect(responseBody.message).toEqual(
      'You attempted to save an item to the database that already exists',
    );
  });

  test('Missing required field and wrong type when attempting to save a book to database', async () => {
    const missingFieldsReq = {
      title: 'Dune',
      author: 5,
      pageCount: 584,
      pubYear: 1965,
      imageUrls: [
        'https://m.media-amazon.com/images/I/81Ua99CURsL._UF894,1000_QL80_.jpg',
      ],
    };
    const req = new NextRequest(`${serverDomain}/database/save`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(missingFieldsReq),
    });

    const res = await POST(req as NextRequest, { params: { type: 'book' } });
    const responseBody: DatabaseSaveEditErrorResponse = await res.json();
    const message = responseBody.message;
    const errors = responseBody.errors;
    expect(message).toEqual('Schema violation(s) during save/edit request');
    expect(errors.length).toEqual(2);
    expect(errors[0]).toEqual('Save/Edit request missing spineColor');
    expect(errors[1]).toEqual('author is of wrong type');

    expect(res.status).toEqual(422);
  });

  test('Attempt database save of book with empty image array', async () => {
    const missingImagesReq = {
      title: 'Title',
      author: 'Author',
      pageCount: 100,
      pubYear: 2025,
      spineColor: 'HexColor',
      imageUrls: [],
    };
    const req = new NextRequest(`${serverDomain}/database/save`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(missingImagesReq),
    });

    const res = await POST(req as NextRequest, { params: { type: 'book' } });
    const responseBody: DatabaseSaveEditErrorResponse = await res.json();
    const message = responseBody.message;
    const errors = responseBody.errors;
    expect(message).toEqual('Schema violation(s) during save/edit request');
    expect(errors[0]).toEqual('Save/Edit request missing imageUrls');
    expect(res.status).toEqual(422);
  });

  test('Attempt pagination retrieval of database items', async () => {
    const req = new NextRequest(
      `${serverDomain}/database?type=book&sort=title&limit=1&page=2&ascDesc=asc`,
      { method: 'GET' },
    );

    const res = await paginationGET(req as NextRequest);
    const responseBody: SuccessfulPaginationResponse = await res.json();
    expect(responseBody.message).toEqual('Successful database gather');
    expect(responseBody.paginatedList[0].title).toEqual('Dune');
  });
});

for (const media of otherMedias) {
  describe(`Test saving ${media}s to database`, () => {
    test(`Attempt database save of ${media} with proper inputs`, async () => {
      const req = new NextRequest(`${serverDomain}/database/save`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(testOtherMediaReq),
      });

      const res = await POST(req as NextRequest, {
        params: { type: `${media}` },
      });
      const responseBody: SuccessfulMediaSaveEditResponse = await res.json();

      const newMedia = responseBody.actionAttemptItem;
      const message = responseBody.message;
      expect(newMedia.id).toBeDefined();
      expect(message).toEqual('Media Title successfully added to database.');
      expect(res.status).toEqual(201);
    });

    test(`Missing required field and wrong type when attempting to save a ${media} to database`, async () => {
      const missingFieldsReq = {
        title: 1234,
        spineColor: '#f25b26',
      };
      const req = new NextRequest(`${serverDomain}/database/save`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(missingFieldsReq),
      });

      const res = await POST(req as NextRequest, {
        params: { type: `${media}` },
      });
      const responseBody: DatabaseSaveEditErrorResponse = await res.json();
      const message = responseBody.message;
      const errors = responseBody.errors;
      expect(message).toEqual('Schema violation(s) during save/edit request');
      expect(errors.length).toEqual(2);
      expect(errors[0]).toEqual('Save/Edit request missing imageUrls');
      expect(errors[1]).toEqual('title is of wrong type');

      expect(res.status).toEqual(422);
    });

    test(`Attempt database save of ${media} with empty image array`, async () => {
      const missingImagesReq = {
        title: 'Title',
        imageUrls: [],
        spineColor: '#f25b26',
      };
      const req = new NextRequest(`${serverDomain}/database/save`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(missingImagesReq),
      });

      const res = await POST(req as NextRequest, {
        params: { type: `${media}` },
      });
      const responseBody: DatabaseSaveEditErrorResponse = await res.json();
      const message = responseBody.message;
      const errors = responseBody.errors;
      expect(message).toEqual('Schema violation(s) during save/edit request');
      expect(errors[0]).toEqual('Save/Edit request missing imageUrls');
      expect(res.status).toEqual(422);
    });
  });
}

describe('Test finding media items by title', () => {
  test('Attempt find book by title, non-common capitalization', async () => {
    const req = new NextRequest(
      `${serverDomain}/database/search?type=book&title=BoOk TiTle`,
      { method: 'GET' },
    );

    const res = await databaseSearchGET(req as NextRequest);
    const body: SuccessfulMediaSearchResponse = await res.json();
    const foundMediaList = body.foundMediaList;
    const message = body.message;
    expect(message).toEqual(
      'Successfully found 1 book(s) with title Book Title',
    );
    expect(res.status).toEqual(200);
    expect(foundMediaList[0].id).toBeDefined();
  });

  test('Attempt find book by title thats not in the database', async () => {
    const req = new NextRequest(
      `${serverDomain}/database/search?type=book&title=DoesNotExist`,
      { method: 'GET' },
    );

    const res = await databaseSearchGET(req as NextRequest);
    const body: SearchErrorResponse = await res.json();
    const message = body.message;
    expect(message).toEqual('No book in database with title DoesNotExist');
    expect(res.status).toEqual(404);
  });

  test('Attempt to search with missing parameters', async () => {
    const req = new NextRequest(`${serverDomain}/database/search?type=book`, {
      method: 'GET',
    });

    const res = await databaseSearchGET(req as NextRequest);
    const responseBody: SearchErrorResponse = await res.json();
    const message = responseBody.message;
    expect(message).toEqual(
      'Search parameters `type` and `title` are required',
    );
    expect(res.status).toEqual(422);
  });
});

describe('Test database deletion', () => {
  const medias: MediaType[] = ['book', 'movie', 'video_game', 'album'];
  for (const media of medias) {
    test(`Delete a ${media} from the database`, async () => {
      let title: string;
      if (media === 'book') {
        title = 'Dune';
      } else {
        title = 'Media Title';
      }

      const req = new NextRequest(
        `${serverDomain}/database/delete?type=${media}&title=${title}`,
        {
          method: 'DELETE',
        },
      );

      const res = await DELETE(req as NextRequest);
      const responseBody: { message: string } = await res.json();
      expect(responseBody.message).toEqual(`Successfully deleted ${title}`);
      expect(res.status).toEqual(200);
    });
  }

  test('Attempt deletion of non-existent item', async () => {
    const title = 'DoesNotExist';
    const media = 'movie';
    const req = new NextRequest(
      `${serverDomain}/database/delete?type=${media}&title=${title}`,
      {
        method: 'DELETE',
      },
    );

    const res = await DELETE(req as NextRequest);
    const responseBody: ErrorResponse = await res.json();
    expect(responseBody.message).toEqual(
      `No item with title:${title} in the ${media} database exists`,
    );
    expect(res.status).toEqual(404);
  });
});

describe('Test edit of database item', () => {
  const medias: MediaType[] = ['book', 'movie', 'video_game', 'album'];
  for (const media of medias) {
    test(`Test edit of ${media}`, async () => {
      const id = saveIDs[`${media}ID`];

      if (media === 'book') {
        bookData.title = 'Book Title Edited';
        bookData.id = id;

        const req = new NextRequest(`${serverDomain}/database/edit`, {
          method: 'PUT',
          body: JSON.stringify(bookData),
        });

        const res = await PUT(req as NextRequest, { params: { type: media } });
        const responseBody: SuccessfulMediaSaveEditResponse = await res.json();

        expect(responseBody.actionAttemptItem.title).toEqual(
          'Book Title Edited',
        );
        expect(responseBody.message).toEqual(
          'Book Title Edited successfully edited.',
        );
      } else {
        otherData.title = 'Other Title Edited';
        otherData.id = id;
        const req = new NextRequest(`${serverDomain}/database/edit`, {
          method: 'PUT',
          body: JSON.stringify(otherData),
        });

        const res = await PUT(req as NextRequest, { params: { type: media } });
        const responseBody: SuccessfulMediaSaveEditResponse = await res.json();

        expect(responseBody.actionAttemptItem.title).toEqual(
          'Other Title Edited',
        );
        expect(responseBody.message).toEqual(
          'Other Title Edited successfully edited.',
        );
      }
    });

    test(`Test edit of non existent media item`, async () => {
      const id = randomUUID();

      if (media === 'book') {
        bookData.title = 'Book Title Edited';
        bookData.id = id;

        const req = new NextRequest(`${serverDomain}/database/edit`, {
          method: 'PUT',
          body: JSON.stringify(bookData),
        });

        const res = await PUT(req as NextRequest, { params: { type: media } });
        const responseBody: DatabaseSaveEditErrorResponse = await res.json();
        expect(responseBody.message).toEqual(
          'Edit requested on an item that does not exist in the database',
        );
        expect(responseBody.errors[0]).toEqual(
          `${bookData.title} does not exist in the database.`,
        );
      } else {
        otherData.title = 'Other Title Edited';
        otherData.id = id;
        const req = new NextRequest(`${serverDomain}/database/edit`, {
          method: 'PUT',
          body: JSON.stringify(otherData),
        });

        const res = await PUT(req as NextRequest, { params: { type: media } });
        const responseBody: DatabaseSaveEditErrorResponse = await res.json();
        expect(responseBody.message).toEqual(
          'Edit requested on an item that does not exist in the database',
        );
        expect(responseBody.errors[0]).toEqual(
          `${otherData.title} does not exist in the database.`,
        );
      }
    });
  }
});

afterAll(async () => {
  // Clean up with Drizzle
  await db.delete(books);
  await db.delete(movies);
  await db.delete(albums);
  await db.delete(videoGames);

  // If you still have a raw pg.Pool, close it in your Jest global teardown
  // where the pool is created, not via the Drizzle `db` instance.
});
