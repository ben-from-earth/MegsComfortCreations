process.env.NODE_ENV = "test";

const request = require("supertest");
const app = require("../app");
const db = require("../database/db");

describe("Connection to database succesful", () => {
  test("can query the database", async () => {
    const result = await db.query("SELECT 1 + 1 AS result");
    expect(result.rows[0].result).toBe(2);
  });
});

const testBookReq = {
  title: "Dune",
  author: "Frank Herbert",
  page_count: 584,
  pub_year: 1965,
  spine_color: "#f25b26",
  image_urls: ["http://testurl.com"],
};

const testOtherMediaReq = {
  title: "Media Title",
  image_urls: ["http://testurl.com"],
};

beforeAll(async function () {
  const data = {
    title: "Ready Player One",
    author: "Ernest Cline",
    page_count: 300,
    pub_year: 2015,
    spine_color: "#ca2f24ff",
    image_urls: ["http://testurl.com"],
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

describe("Test saving book to database", () => {
  test("Attempt database save with proper inputs", async () => {
    const res = await request(app)
      .post("/database/save/book")
      .send(testBookReq);
    const newBook = res.body.saved_book;
    const message = res.body.message;
    expect(newBook.id).toBeDefined();
    expect(message).toEqual("Dune successfully added to database.");
    expect(res.statusCode).toEqual(201);
  });

  test("Attempt database save of existing title/author combo", async () => {
    const res = await request(app)
      .post("/database/save/book")
      .send(testBookReq);
    const error = res.body.error;
    const message = res.body.message;
    expect(error).toEqual("Error saving book to database");
    expect(message).toEqual(
      "Key (title, author)=(Dune, Frank Herbert) already exists."
    );
    expect(res.statusCode).toEqual(400);
  });

  test("Missing required field and wrong type when attempting to save a book to database", async () => {
    const missingFieldsReq = {
      title: "Dune",
      author: 5,
      page_count: 584,
      pub_year: 1965,
      image_urls: [
        "https://m.media-amazon.com/images/I/81Ua99CURsL._UF894,1000_QL80_.jpg",
      ],
    };
    const res = await request(app)
      .post("/database/save/book")
      .send(missingFieldsReq);
    const error = res.body.error;
    const validationErrors = res.body.validationErrors;
    expect(error).toEqual("Schema violation during save request");
    expect(validationErrors.length).toEqual(2);
    expect(validationErrors[0]).toEqual(
      "Database save request is missing field: spine_color"
    );
    expect(validationErrors[1]).toEqual(
      "author was input as the wrong type, should be a(n) string"
    );

    expect(res.statusCode).toEqual(400);
  });

  test("Attempt database save of book with empty image array", async () => {
    const missingImagesReq = {
      title: "Title",
      author: "Author",
      page_count: 100,
      pub_year: 2025,
      spine_color: "HexColor",
      image_urls: [],
    };
    const res = await request(app)
      .post("/database/save/book")
      .send(missingImagesReq);
    const error = res.body.error;
    const validationErrors = res.body.validationErrors;
    expect(error).toEqual("Schema violation during save request");
    expect(validationErrors[0]).toEqual(
      "Database save attempted without images"
    );
    expect(res.statusCode).toEqual(400);
  });

  test("Attempt pagination retrieval of database items", async () => {
    const res = await request(app).get(
      "/database?type=book&sort=title&limit=1&page=2"
    );
    expect(res.body.message).toEqual("Successful database gather");
    expect(res.body.paginatedList[0].title).toEqual("Ready Player One");
  });
});

describe("Test saving other media to database", () => {
  const otherMedias = ["movie", "album", "video_game"];
  for (let media of otherMedias) {
    test(`Attempt database save of ${media} with proper inputs`, async () => {
      const res = await request(app)
        .post(`/database/save/${media}`)
        .send(testOtherMediaReq);
      const newMedia = res.body[`saved_${media}`];
      const message = res.body.message;
      expect(newMedia.id).toBeDefined();
      expect(message).toEqual("Media Title successfully added to database.");
      expect(res.statusCode).toEqual(201);
    });

    test(`Missing required field and wrong type when attempting to save a ${media} to database`, async () => {
      const missingFieldsReq = {
        title: 1234,
      };
      const res = await request(app)
        .post(`/database/save/${media}`)
        .send(missingFieldsReq);
      const error = res.body.error;
      const validationErrors = res.body.validationErrors;
      expect(error).toEqual("Schema violation during save request");
      expect(validationErrors.length).toEqual(2);
      expect(validationErrors[0]).toEqual(
        "Database save request is missing field: image_urls"
      );
      expect(validationErrors[1]).toEqual(
        "title was input as the wrong type, should be a(n) string"
      );

      expect(res.statusCode).toEqual(400);
    });
  }
});

describe("Test finding media items by title", () => {
  test("Attempt find book by title, non-common capitalization", async () => {
    const res = await request(app).get(
      "/database/search?type=book&title=ReADy PlaYeR ONE"
    );
    const foundBooksList = res.body.foundBooksList;
    const message = res.body.message;
    expect(message).toEqual(
      "Successfully found 1 book(s) with title Ready Player One"
    );
    expect(res.statusCode).toEqual(200);
    expect(foundBooksList[0].id).toBeDefined();
  });

  test("Attempt find book by title thats not in the database", async () => {
    const res = await request(app).get(
      "/database/search?type=book&title=DoesNotExist"
    );
    const message = res.body.message;
    expect(message).toEqual("No book in database with title DoesNotExist");
    expect(res.statusCode).toEqual(404);
  });
});

afterAll(async function () {
  await db.query("DELETE FROM books");
  await db.query("DELETE FROM movies");
  await db.query("DELETE FROM albums");
  await db.query("DELETE FROM video_games");
  await db.end();
});
