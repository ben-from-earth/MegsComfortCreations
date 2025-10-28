import db from '@/lib/database/db';
import {
  MediaType,
  postSavedMediaItem,
  presavedMediaItem,
  SuccessfulMediaSaveEditResponse,
  SuccessfulMediaSearchResponse,
  SuccessfulPaginationResponse,
} from '@/lib/interfaces/globalInterfaces';
import { GET as databaseSearchGET } from '@/app/api/database/search/route';
import { GET as paginationGET } from '@/app/api/database/route';
import { DELETE } from '@/app/api/database/delete/route';
import { NextRequest } from 'next/server';
import {
  ApiError,
  DatabaseSaveEditErrorResponse,
  ErrorResponse,
  SearchErrorResponse,
} from '../app/api/api-Errors';
import { POST } from '@/app/api/database/save/[type]/route';
import { PUT } from '@/app/api/database/edit/[type]/route';
import { randomUUID } from 'crypto';

const serverDomain = process.env.SERVER_BASE_URL;

describe('Connection to database succesful', () => {
  test('can query the database', async () => {
    const result = await db.query<{ result: number }>('SELECT 1 + 1 AS result');
    expect(result.rows[0].result).toBe(2);
  });
});

const otherMedias: MediaType[] = ['movie', 'video_game', 'album'];
const saveIDs = {
  bookID: '',
  movieID: '',
  video_gameID: '',
  albumID: '',
};

const testBookReq = {
  title: 'Dune',
  author: 'Frank Herbert',
  page_count: 584,
  pub_year: 1965,
  spine_color: '#f25b26',
  image_urls: ['http://testurl.com'],
};

const testOtherMediaReq = {
  title: 'Media Title',
  spine_color: '#f25b26',
  image_urls: ['http://testurl.com'],
};

const bookData: presavedMediaItem & { id?: string } = {
  title: 'Book Title',
  author: 'Book Author',
  page_count: 100,
  pub_year: 2025,
  spine_color: '#ca2f24ff',
  image_urls: ['http://testurl.com'],
};
const otherData: presavedMediaItem & { id?: string } = {
  title: 'Other Title',
  spine_color: '#ca2f24ff',
  image_urls: ['http://testurl.com'],
};

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
    ],
  );

  saveIDs.bookID = bookResult.rows[0].id;

  const movieResult = await db.query<postSavedMediaItem>(
    `INSERT INTO movies (
            title,
            spine_color,
            image_urls 
      ) VALUES ($1, $2, $3) 
      RETURNING 
            id,
            title,
            spine_color,
            image_urls`,
    [otherData.title, otherData.spine_color, otherData.image_urls],
  );

  saveIDs.movieID = movieResult.rows[0].id;

  const VGResult = await db.query<postSavedMediaItem>(
    `INSERT INTO video_games (
            title,
            image_urls,
            spine_color 
      ) VALUES ($1, $2, $3) 
      RETURNING 
            id,
            image_urls,
            spine_color`,
    [otherData.title, otherData.image_urls, otherData.spine_color],
  );

  saveIDs.video_gameID = VGResult.rows[0].id;

  const albumResult = await db.query<postSavedMediaItem>(
    `INSERT INTO albums (
            title,
            image_urls,
            spine_color 
      ) VALUES ($1, $2, $3) 
      RETURNING 
            id,
            image_urls,
            spine_color`,
    [otherData.title, otherData.image_urls, otherData.spine_color],
  );

  saveIDs.albumID = albumResult.rows[0].id;
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
      page_count: 584,
      pub_year: 1965,
      image_urls: [
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
    expect(errors[0]).toEqual('Save/Edit request missing spine_color');
    expect(errors[1]).toEqual('author is of wrong type');

    expect(res.status).toEqual(422);
  });

  test('Attempt database save of book with empty image array', async () => {
    const missingImagesReq = {
      title: 'Title',
      author: 'Author',
      page_count: 100,
      pub_year: 2025,
      spine_color: 'HexColor',
      image_urls: [],
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
    expect(errors[0]).toEqual('Save/Edit request missing image_urls');
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
        spine_color: '#f25b26',
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
      expect(errors[0]).toEqual('Save/Edit request missing image_urls');
      expect(errors[1]).toEqual('title is of wrong type');

      expect(res.status).toEqual(422);
    });

    test(`Attempt database save of ${media} with empty image array`, async () => {
      const missingImagesReq = {
        title: 'Title',
        image_urls: [],
        spine_color: '#f25b26',
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
      expect(errors[0]).toEqual('Save/Edit request missing image_urls');
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
  for (let media of medias) {
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

  //test deletion of an item that has a duplicate title in the database
});

describe('Test edit of database item', () => {
  const medias: MediaType[] = ['book', 'movie', 'video_game', 'album'];
  for (let media of medias) {
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

afterAll(async function () {
  await db.query('DELETE FROM books');
  await db.query('DELETE FROM movies');
  await db.query('DELETE FROM albums');
  await db.query('DELETE FROM video_games');
  await db.end();
});
