process.env.NODE_ENV = 'test';

const request = require('supertest');
const app = require('../app');
const db = require('../database/db');

describe('Connection to database succesful', () => {
  test('can query the database', async () => {
    const result = await db.query('SELECT 1 + 1 AS result');
    expect(result.rows[0].result).toBe(2);
  });
});

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
  image_urls: ['http://testurl.com'],
};

beforeAll(async function () {
  const data = {
    title: 'Ready Player One',
    author: 'Ernest Cline',
    page_count: 300,
    pub_year: 2015,
    spine_color: '#ca2f24ff',
    image_urls: ['http://testurl.com'],
  };
  const result = await db.query(
    `INSERT INTO books (
            title,
            author,
            page_count,
            pub_year,
            image_urls,
            spine_color 
      ) VALUES ($1, $2, $3, $4, $5, $6) 
      RETURNING 
            id,
            title,
            author,
            page_count,
            pub_year,
            image_urls,
            spine_color`,
    [
      data.title,
      data.author,
      data.page_count,
      data.pub_year,
      data.image_urls,
      data.spine_color,
    ]
  );
});

describe('Test saving book to database', () => {
  test('Attempt database save with proper inputs', async () => {
    const res = await request(app)
      .post('/database/save/book')
      .send(testBookReq);
    const newBook = res.body.saveAttemptItem;
    const message = res.body.message;
    expect(newBook.id).toBeDefined();
    expect(message).toEqual('Dune successfully added to database.');
    expect(res.statusCode).toEqual(201);
  });

  test('Attempt database save of existing title/author combo', async () => {
    const res = await request(app)
      .post('/database/save/book')
      .send(testBookReq);
    const saved = res.body.saved;
    const message = res.body.message;
    expect(saved).toEqual(false);
    expect(message).toEqual(
      'Key (title, author)=(Dune, Frank Herbert) already exists.'
    );
    expect(res.statusCode).toEqual(400);
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
    const res = await request(app)
      .post('/database/save/book')
      .send(missingFieldsReq);
    const message = res.body.message;
    const errors = res.body.errors;
    expect(message).toEqual('Schema violation(s) during save request');
    expect(errors.length).toEqual(2);
    expect(errors[0]).toEqual('Save request missing spine_color');
    expect(errors[1]).toEqual('author is of wrong type');

    expect(res.statusCode).toEqual(400);
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
    const res = await request(app)
      .post('/database/save/book')
      .send(missingImagesReq);
    const message = res.body.message;
    const errors = res.body.errors;
    expect(message).toEqual('Schema violation(s) during save request');
    expect(errors[0]).toEqual('Save request missing image_urls');
    expect(res.statusCode).toEqual(400);
  });

  test('Attempt pagination retrieval of database items', async () => {
    const res = await request(app).get(
      '/database?type=book&sort=title&limit=1&page=2'
    );
    expect(res.body.message).toEqual('Successful database gather');
    expect(res.body.paginatedList[0].title).toEqual('Ready Player One');
  });
});
const otherMedias = ['movie', 'album', 'video_game'];
for (let media of otherMedias) {
  describe(`Test saving ${media}s to database`, () => {
    test(`Attempt database save of ${media} with proper inputs`, async () => {
      const res = await request(app)
        .post(`/database/save/${media}`)
        .send(testOtherMediaReq);
      const newMedia = res.body.saveAttemptItem;
      const message = res.body.message;
      expect(newMedia.id).toBeDefined();
      expect(message).toEqual('Media Title successfully added to database.');
      expect(res.statusCode).toEqual(201);
    });

    test(`Missing required field and wrong type when attempting to save a ${media} to database`, async () => {
      const missingFieldsReq = {
        title: 1234,
      };
      const res = await request(app)
        .post(`/database/save/${media}`)
        .send(missingFieldsReq);
      const message = res.body.message;
      const errors = res.body.errors;
      expect(message).toEqual('Schema violation(s) during save request');
      expect(errors.length).toEqual(2);
      expect(errors[0]).toEqual('Save request missing image_urls');
      expect(errors[1]).toEqual('title is of wrong type');

      expect(res.statusCode).toEqual(400);
    });

    test(`Attempt database save of ${media} with empty image array`, async () => {
      const missingImagesReq = {
        title: 'Title',
        image_urls: [],
      };
      const res = await request(app)
        .post(`/database/save/${media}`)
        .send(missingImagesReq);
      const message = res.body.message;
      const errors = res.body.errors;
      expect(message).toEqual('Schema violation(s) during save request');
      expect(errors[0]).toEqual('Save request missing image_urls');
      expect(res.statusCode).toEqual(400);
    });
  });
}

describe('Test finding media items by title', () => {
  test('Attempt find book by title, non-common capitalization', async () => {
    const res = await request(app).get(
      '/database/search?type=book&title=ReADy PlaYeR ONE'
    );
    const foundMediaList = res.body.foundMediaList;
    const message = res.body.message;
    expect(message).toEqual(
      'Successfully found 1 book(s) with title Ready Player One'
    );
    expect(res.statusCode).toEqual(200);
    expect(foundMediaList[0].id).toBeDefined();
  });

  test('Attempt find book by title thats not in the database', async () => {
    const res = await request(app).get(
      '/database/search?type=book&title=DoesNotExist'
    );
    const message = res.body.message;
    expect(message).toEqual('No book in database with title DoesNotExist');
    expect(res.statusCode).toEqual(404);
  });
});

describe('Test database deletion', () => {
  const medias = ['book', 'movie', 'video_game', 'album'];
  for (let media of medias) {
    test(`Delete a ${media} from the database`, async () => {
      let title;
      if (media === 'book') {
        title = 'Dune';
      } else {
        title = 'Media Title';
      }
      const res = await request(app).delete(
        `/database?type=${media}&title=${title}`
      );
      const message = res.body.message;
      expect(message).toEqual(`Successfully deleted ${title}`);
      expect(res.statusCode).toEqual(200);
    });
  }

  test('Attempt deletion of non-existent item', async () => {
    const title = 'DoesNotExist';
    const media = 'movie';
    const res = await request(app).delete(
      `/database?type=${media}&title=${title}`
    );
    const message = res.body.message;
    expect(message).toEqual(
      `No item with title:${title} in the ${media} database exists`
    );
    expect(res.statusCode).toEqual(400);
  });
});

afterAll(async function () {
  await db.query('DELETE FROM books');
  await db.query('DELETE FROM movies');
  await db.query('DELETE FROM albums');
  await db.query('DELETE FROM video_games');
  await db.end();
});
