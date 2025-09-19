const express = require('express');
const router = new express.Router();
const { validate } = require('jsonschema');

//get Models
const Book = require('../models/book');
const Movie = require('../models/movie');
const Video_Game = require('../models/video_game');
const Album = require('../models/album');

//get Schemas
const bookCreateSchema = require('../schemas/bookCreateSchema.json');
const otherMediaCreateSchema = require('../schemas/otherMediaCreateSchema.json');

//custom Error
const db = require('../database/db');

const { titleRearrange } = require('../helpers/mediaCollectorHelpers');

router.post('/save/:type', async (req, res, next) => {
  switch (req.params.type) {
    case 'book':
      try {
        const validation = validate(req.body, bookCreateSchema);
        if (!validation.valid) {
          return next({
            status: 400,
            schemaErrors: validation.errors.map((e) => e.stack),
            saveAttemptItem: req.body,
            type: req.params.type,
          });
        }
        const book = await Book.create(req.body);
        //book.id will give in database id
        return res.status(201).json({
          message: `${titleRearrange(
            req.body.title
          )} successfully added to database.`,
          saveAttemptItem: book,
          saved: true,
          type: req.params.type,
        });
      } catch (err) {
        const error = {
          ...err,
          saveAttemptItem: req.body,
          type: req.params.type,
        };
        return next(error);
      }

    case 'movie':
      try {
        const validation = validate(req.body, otherMediaCreateSchema);
        if (!validation.valid) {
          return next({
            status: 400,
            schemaErrors: validation.errors.map((e) => e.stack),
            saveAttemptItem: req.body,
            type: req.params.type,
          });
        }
        const movie = await Movie.create(req.body);
        return res.status(201).json({
          message: `${titleRearrange(
            req.body.title
          )} successfully added to database.`,
          saveAttemptItem: movie,
          saved: true,
          type: req.params.type,
        });
      } catch (err) {
        return next(err);
      }
    case 'video_game':
      try {
        const validation = validate(req.body, otherMediaCreateSchema);
        if (!validation.valid) {
          return next({
            status: 400,
            schemaErrors: validation.errors.map((e) => e.stack),
            saveAttemptItem: req.body,
            type: req.params.type,
          });
        }
        const video_game = await Video_Game.create(req.body);
        return res.status(201).json({
          message: `${titleRearrange(
            req.body.title
          )} successfully added to database.`,
          saveAttemptItem: video_game,
          saved: true,
          type: req.params.type,
        });
      } catch (err) {
        return next(err);
      }
    case 'album':
      try {
        const validation = validate(req.body, otherMediaCreateSchema);
        if (!validation.valid) {
          return next({
            status: 400,
            schemaErrors: validation.errors.map((e) => e.stack),
            saveAttemptItem: req.body,
            type: req.params.type,
          });
        }
        const album = await Album.create(req.body);
        return res.status(201).json({
          message: `${titleRearrange(
            req.body.title
          )} successfully added to database.`,
          saveAttemptItem: album,
          saved: true,
          type: req.params.type,
        });
      } catch (err) {
        return next(err);
      }
  }
});

router.get('/search', async (req, res, next) => {
  switch (req.query.type) {
    case 'book':
      try {
        const bookList = await Book.find(req.query.title);
        if (bookList.length === 0) {
          return next({
            status: 404,
            error: 'Media not found',
            message: `No book in database with title ${req.query.title}`,
          });
        }
        return res.status(200).json({
          message: `Successfully found ${bookList.length} book(s) with title ${bookList[0].title}`,
          foundMediaList: bookList,
        });
      } catch (err) {
        return next(err);
      }
    case 'movie':
      try {
        const movieList = await Movie.find(req.query.title);
        if (movieList.length === 0) {
          return next({
            status: 404,
            error: 'Media not found',
            message: `No movie in database with title ${req.query.title}`,
          });
        }
        return res.status(200).json({
          message: `Successfully found ${movieList.length} movie(s) with title ${movieList[0].title}`,
          foundMediaList: movieList,
        });
      } catch (err) {
        return next(err);
      }
    case 'video_game':
      try {
        const VGList = await Video_Game.find(req.query.title);
        if (VGList.length === 0) {
          return next({
            status: 404,
            error: 'Media not found',
            message: `No video game in database with title ${req.query.title}`,
          });
        }
        return res.status(200).json({
          message: `Successfully found ${VGList.length} video game(s) with title ${VGList[0].title}`,
          foundMediaList: VGList,
        });
      } catch (err) {
        return next(err);
      }
    case 'album':
      try {
        const albumList = await Album.find(req.query.title);
        if (albumList.length === 0) {
          return next({
            status: 404,
            error: 'Media not found',
            message: `No album in database with title ${req.query.title}`,
          });
        }
        return res.status(200).json({
          message: `Successfully found ${albumList.length} album(s) with title ${albumList[0].title}`,
          foundMediaList: albumList,
        });
      } catch (err) {
        return next(err);
      }
  }
});

router.get('/', async (req, res, next) => {
  // /database?type=movie&limit=5&page=2
  // SELECT * FROM movies ORDER BY title LIMIT 5 OFFSET 5

  //All of these options are handled by the front end so errors will be prevented before the request.
  const limit = Number(req.query.limit);
  const page = Number(req.query.page) || 1;
  const type = req.query.type;
  const sort = req.query.sort;

  const offset = (page - 1) * limit;
  try {
    const result = await db.query(
      `SELECT * 
          FROM ${type + 's'}
          ORDER BY ${sort}
          LIMIT $1 OFFSET $2`,
      [limit, offset]
    );
    paginatedList = result.rows;

    const totalRes = await db.query(
      `SELECT COUNT(*) 
          FROM ${type + 's'}`
    );

    const total = parseInt(totalRes.rows[0].count, 10);
    return res.status(200).json({
      message: `Successful database gather`,
      paginatedList,
      total,
    });
  } catch (error) {
    return next({
      status: 400,
      error: 'Database collection error',
      message: 'Error gathering items from the database during pagination',
    });
  }
});

router.get('/titleSearch', async (req, res, next) => {
  // /database/titleSearch?type=movie&title=Finding Nemo
  // SELECT * FROM movies WHERE title ILIKE $1 || '%' ORDER BY title

  //All of these options are handled by the front end so errors will be prevented before the request.
  const type = req.query.type;
  const title = req.query.title;
  try {
    const result = await db.query(
      `SELECT * 
          FROM ${type + 's'}
          WHERE title ILIKE $1 || '%'
          ORDER BY title`,
      [title]
    );
    const titleSearchResponse = result.rows;
    return res.status(200).json({
      message: `Successful database gather`,
      titleSearchResponse,
      total: titleSearchResponse.length,
    });
  } catch (error) {
    return next({
      status: 400,
      error: 'Database collection error',
      message: 'Error gathering items from the database during search',
    });
  }
});

router.delete('/', async (req, res, next) => {
  const type = req.query.type;
  const title = req.query.title;

  try {
    const deleteRes = await db.query(
      `DELETE FROM ${type + 's'}
          WHERE title ILIKE $1`,
      [title]
    );

    if (deleteRes.rowCount === 0) {
      next({
        status: 400,
        error: 'Non-existent deletion request',
        message: `No item with title:${title} in the ${type} database exists`,
      });
    } else {
      return res.status(200).json({
        message: `Successfully deleted ${title}`,
      });
    }
  } catch (error) {
    return next({
      status: 400,
      error: 'Database deletion error',
      message: 'Error deleting items from the database.',
    });
  }
});

module.exports = router;
