process.env.NODE_ENV = 'test';

const request = require('supertest');
const app = require('../app');
const db = require('../database/db');
const { randomUUID } = require('crypto');

describe('Connection to database succesful', () => {
  test('can query the database', async () => {
    const result = await db.query('SELECT 1 + 1 AS result');
    expect(result.rows[0].result).toBe(2);
  });
});

const bookID = randomUUID();

describe('Link/unlink book to genre', () => {
  test('Link bookID with Science Fiction and Fantasy', async () => {
    const res = await request(app)
      .post('/genres/addLink')
      .send({ bookID, genres: ['Science Fiction', 'Fantasy'] });
    expect(res.body.responses.length).toEqual(2);
    expect(res.body.responses[1].genre).toEqual('Fantasy');
  });

  test('Unlink Fantasy from book', async () => {
    const res = await request(app)
      .post('/genres/unlink')
      .send({ bookID, genres: ['Fantasy'] });
    expect(res.body.responses.length).toEqual(1);
    expect(res.body.responses[0].message).toEqual(
      'Successfully removed genre: Fantasy'
    );
  });
});

describe('Get genres requests', () => {
  test('Get all genres', async () => {
    const res = await request(app).get('/genres/getAll');

    expect(res.body.genres).toBeDefined();
  });

  test('Get genres for bookID', async () => {
    const res = await request(app).get(`/genres/getForBook?bookID=${bookID}`);
    expect(res.body.message).toEqual(
      `Successfully grabbed genres for bookID ${bookID}`
    );
    expect(res.body.genres.length).toEqual(1);
    expect(res.body.genres[0]).toEqual('Science Fiction');
  });

  test('Get all books with Science Fiction as a genre', async () => {
    const testBookReq = {
      title: 'Genre Test',
      author: 'author',
      page_count: 100,
      pub_year: 2025,
      spine_color: '#f25b26',
      image_urls: ['http://testurl.com'],
    };
    //add book to database with Science Fiction link
    const bookSaveRes = await request(app)
      .post('/database/save/book')
      .send(testBookReq);
    const newBookID = bookSaveRes.body.saveAttemptItem.id;

    const scifiLinkRes = await request(app)
      .post('/genres/addLink')
      .send({ bookID: newBookID, genres: ['Science Fiction'] });

    const res = await request(app).get(
      '/genres?genre=Science Fiction&sort=title&limit=3&page=1&ascDesc=asc'
    );
    expect(res.body.message).toEqual(`Successful database gather`);
    expect(res.body.total).toEqual(1);
    expect(res.body.paginatedList[0].title).toEqual('Genre Test');
  });
});

describe('Handle Proper delete of a book', () => {
  test('Delete all links for a given bookID', async () => {
    const res = await request(app).get(
      `/genres/removeAllLinksForBook?bookID=${bookID}`
    );
    expect(res.body.message).toEqual(
      `All genre links removed for bookID ${bookID}`
    );
    expect(res.body.actionCompleted).toEqual(true);

    const checkRes = await request(app).get(
      `/genres/getForBook?bookID=${bookID}`
    );
    expect(checkRes.body.genres.length).toEqual(0);
  });
});

afterAll(async function () {
  await db.query('DELETE FROM genres_books');
  await db.query('DELETE FROM books');
  await db.end();
});
