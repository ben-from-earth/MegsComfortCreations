const express = require('express');
const router = new express.Router();

const Genre = require('../models/genre');
const db = require('../database/db');

router.get('/getAll', async (req, res, next) => {
  try {
    const genres = await Genre.getAllGenres();
    return res.status(200).json({ message: 'success', genres });
  } catch (err) {
    return next({
      status: 400,
      error: 'Genre Error',
      message: 'Error connecting to the database and/or genre table',
    });
  }
});

router.get('/getForBook', async (req, res, next) => {
  try {
    const bookID = req.query.bookID;
    const genres = await Genre.getForBook(bookID);

    return res.status(200).json({
      message: `Successfully grabbed genres for bookID ${bookID}`,
      genres,
    });
  } catch (err) {
    return next({
      status: 400,
      error: 'Genre Error',
      message: 'Error connecting to the database and/or genre table',
    });
  }
});

router.post('/addlink', async (req, res, next) => {
  const bookID = req.body.bookID;
  const genres = req.body.genres;
  let responses = [];
  for (let genre of genres) {
    try {
      await Genre.link(genre, bookID);
      responses.push({ message: 'Successful genre link', genre, bookID });
    } catch (err) {
      return next({
        status: 400,
        error: 'Genre Error',
        message: 'Error connecting to the database and/or genre table',
      });
    }
  }
  return res.status(201).json({ responses });
});

router.post('/unlink', async (req, res, next) => {
  const bookID = req.body.bookID;
  const genres = req.body.genres;
  let responses = [];
  for (let genre of genres) {
    try {
      await Genre.unlink(bookID, genre);
      responses.push({ message: `Successfully removed genre: ${genre}` });
    } catch (err) {
      return next({
        status: 400,
        error: 'Genre Error',
        message: 'Error connecting to the database and/or genre table',
      });
    }
  }
  return res.status(200).json({ responses });
});

router.get('/', async (req, res, next) => {
  const genre = req.query.genre;

  //All of these options are handled by the front end so errors will be prevented before the request.
  const limit = Number(req.query.limit);
  const page = Number(req.query.page) || 1;
  const sort = req.query.sort;
  const ascDesc = req.query.ascDesc;

  const offset = (page - 1) * limit;
  try {
    const genreRes = await Genre.getBooksWithGenre(
      genre,
      sort,
      offset,
      limit,
      ascDesc
    );
    return res.status(200).json({
      message: `Successful database gather`,
      paginatedList: genreRes.books,
      total: genreRes.total,
    });
  } catch (err) {
    return next({
      status: 400,
      error: 'Genre Error',
      message: 'Error connecting to the database and/or genre table',
    });
  }
});
router.get('/noGenres', async (req, res, next) => {
  const limit = Number(req.query.limit);
  const page = Number(req.query.page) || 1;
  const sort = req.query.sort;
  const ascDesc = req.query.ascDesc;

  const offset = (page - 1) * limit;
  try {
    const genreRes = await Genre.getNoGenreBooks(sort, offset, limit, ascDesc);
    return res.status(200).json({
      message: `Successful database gather`,
      paginatedList: genreRes.books,
      total: genreRes.total,
    });
  } catch (err) {
    return next({
      status: 400,
      error: 'Genre Error',
      message: 'Error connecting to the database and/or genre table',
    });
  }
});

module.exports = router;
