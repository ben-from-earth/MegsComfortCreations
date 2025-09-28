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

const otherMedias = ['movie', 'video_game', 'album'];
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

const bookData = {
  title: 'Book Title',
  author: 'Book Author',
  page_count: 100,
  pub_year: 2025,
  spine_color: '#ca2f24ff',
  image_urls: ['http://testurl.com'],
};
const otherData = {
  title: 'Other Title',
  spine_color: '#ca2f24ff',
  image_urls: ['http://testurl.com'],
};

beforeAll(async function () {
  const bookResult = await db.query(
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

  saveIDs.bookID = bookResult.rows[0].id;

  const movieResult = await db.query(
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
    [otherData.title, otherData.spine_color, otherData.image_urls]
  );

  saveIDs.movieID = movieResult.rows[0].id;

  const VGResult = await db.query(
    `INSERT INTO video_games (
            title,
            image_urls,
            spine_color 
      ) VALUES ($1, $2, $3) 
      RETURNING 
            id,
            image_urls,
            spine_color`,
    [otherData.title, otherData.image_urls, otherData.spine_color]
  );

  saveIDs.video_gameID = VGResult.rows[0].id;

  const albumResult = await db.query(
    `INSERT INTO albums (
            title,
            image_urls
      ) VALUES ($1, $2) 
      RETURNING 
            id,
            image_urls`,
    [otherData.title, otherData.image_urls]
  );

  saveIDs.albumID = albumResult.rows[0].id;
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

  test('Attempt 100 database saves of existing title/author combo', async () => {
    const responses = [];
    for (let i = 0; i <= 100; i++) {
      const res = await request(app)
        .post('/database/save/book')
        .send(testBookReq);
      responses.push(res.body);
    }
    expect(
      responses
        .map((response) => response.actionCompleted)
        .every((ac) => ac === false)
    ).toEqual(true);
    expect(
      responses
        .map((response) => response.message)
        .every(
          (message) =>
            message ===
            'Key (title, author)=(Dune, Frank Herbert) already exists.'
        )
    ).toEqual(true);
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
    expect(message).toEqual('Schema violation(s) during save/edit request');
    expect(errors.length).toEqual(2);
    expect(errors[0]).toEqual('Save/Edit request missing spine_color');
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
    expect(message).toEqual('Schema violation(s) during save/edit request');
    expect(errors[0]).toEqual('Save/Edit request missing image_urls');
    expect(res.statusCode).toEqual(400);
  });

  test('Attempt pagination retrieval of database items', async () => {
    const res = await request(app).get(
      '/database?type=book&sort=title&limit=1&page=2&ascDesc=asc'
    );
    expect(res.body.message).toEqual('Successful database gather');
    expect(res.body.paginatedList[0].title).toEqual('Dune');
  });
});

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
        spine_color: '#f25b26',
      };
      const res = await request(app)
        .post(`/database/save/${media}`)
        .send(missingFieldsReq);
      const message = res.body.message;
      const errors = res.body.errors;
      expect(message).toEqual('Schema violation(s) during save/edit request');
      expect(errors.length).toEqual(2);
      expect(errors[0]).toEqual('Save/Edit request missing image_urls');
      expect(errors[1]).toEqual('title is of wrong type');

      expect(res.statusCode).toEqual(400);
    });

    test(`Attempt database save of ${media} with empty image array`, async () => {
      const missingImagesReq = {
        title: 'Title',
        image_urls: [],
        spine_color: '#f25b26',
      };
      const res = await request(app)
        .post(`/database/save/${media}`)
        .send(missingImagesReq);
      const message = res.body.message;
      const errors = res.body.errors;
      expect(message).toEqual('Schema violation(s) during save/edit request');
      expect(errors[0]).toEqual('Save/Edit request missing image_urls');
      expect(res.statusCode).toEqual(400);
    });
  });
}

describe('Test finding media items by title', () => {
  test('Attempt find book by title, non-common capitalization', async () => {
    const res = await request(app).get(
      '/database/search?type=book&title=BoOk TiTle'
    );
    const foundMediaList = res.body.foundMediaList;
    const message = res.body.message;
    expect(message).toEqual(
      'Successfully found 1 book(s) with title Book Title'
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
    expect(res.statusCode).toEqual(404);
  });
});

describe('Test edit of database item', () => {
  const medias = ['book', 'movie', 'video_game', 'album'];
  for (let media of medias) {
    test(`Test edit of ${media}`, async () => {
      const id = saveIDs[`${media}ID`];
      if (media === 'book') {
        bookData.title = 'Book Title Edited';
        bookData.id = id;
        const bookRes = await request(app)
          .put(`/database/edit/${media}`)
          .send(bookData);
        expect(bookRes.body.editAttemptItem.title).toEqual('Book Title Edited');
        expect(bookRes.body.actionCompleted).toEqual(true);
      } else {
        otherData.title = 'Other Title Edited';
        otherData.id = id;
        const otherRes = await request(app)
          .put(`/database/edit/${media}`)
          .send(otherData);
        expect(otherRes.body.editAttemptItem.title).toEqual(
          'Other Title Edited'
        );
        expect(otherRes.body.actionCompleted).toEqual(true);
      }
    });

    test(`Test edit of non existent media item`, async () => {
      const id = randomUUID();
      if (media === 'book') {
        bookData.title = 'Book Title Edited';
        bookData.id = id;
        const bookRes = await request(app)
          .put(`/database/edit/${media}`)
          .send(bookData);
        expect(bookRes.body.message).toEqual(
          'Edit requested on an item that does not exist in the database'
        );
        expect(bookRes.body.actionCompleted).toEqual(false);
      } else {
        otherData.title = 'Other Title Edited';
        otherData.id = id;
        const otherRes = await request(app)
          .put(`/database/edit/${media}`)
          .send(otherData);
        expect(otherRes.body.message).toEqual(
          'Edit requested on an item that does not exist in the database'
        );
        expect(otherRes.body.actionCompleted).toEqual(false);
      }
    });

    test(`Test edit with improper inputs`, async () => {
      const id = saveIDs[`${media}ID`];
      if (media === 'book') {
        bookData.title = 123;
        bookData.id = id;
        const bookRes = await request(app)
          .put(`/database/edit/${media}`)
          .send(bookData);
        expect(bookRes.body.message).toEqual(
          'Schema violation(s) during save/edit request'
        );
        //possibly update to handle this
        //expect(bookRes.body.actionCompleted).toEqual(false);
        expect(bookRes.body.errors[0]).toEqual(`title is of wrong type`);
      } else {
        otherData.title = 123;
        otherData.id = id;
        const otherRes = await request(app)
          .put(`/database/edit/${media}`)
          .send(otherData);
        expect(otherRes.body.message).toEqual(
          'Schema violation(s) during save/edit request'
        );
        //possibly update to handle this
        // expect(otherRes.body.actionCompleted).toEqual(false);
        expect(otherRes.body.errors[0]).toEqual(`title is of wrong type`);
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
