const express = require("express");
const router = new express.Router();
const { validate } = require("jsonschema");

//get Models
const Book = require("../models/book");
const Movie = require("../models/movie");
const Video_Game = require("../models/video_game");
const Album = require("../models/album");

//get Schemas
const bookCreateSchema = require("../schemas/bookCreateSchema.json");
const otherMediaCreateSchema = require("../schemas/otherMediaCreateSchema.json");

//custom Error
const ExpressError = require("../expressError");

router.post("/save/:type", async (req, res, next) => {
  switch (req.params.type) {
    case "book":
      try {
        const validation = validate(req.body, bookCreateSchema);
        if (!validation.valid) {
          return next({
            status: 400,
            schemaErrors: validation.errors.map((e) => e.stack),
          });
        }
        const book = await Book.create(req.body);
        //book.id will give in database id
        return res.status(201).json({
          message: `${req.body.title} successfully added to database.`,
          saved_book: book,
        });
      } catch (err) {
        return next(err);
      }
    case "movie":
      try {
        const validation = validate(req.body, otherMediaCreateSchema);
        if (!validation.valid) {
          return next({
            status: 400,
            schemaErrors: validation.errors.map((e) => e.stack),
          });
        }
        const movie = await Movie.create(req.body);
        return res.status(201).json({
          message: `${req.body.title} successfully added to database.`,
          saved_movie: movie,
        });
      } catch (err) {
        return next(err);
      }
    case "video_game":
      try {
        const validation = validate(req.body, otherMediaCreateSchema);
        if (!validation.valid) {
          return next({
            status: 400,
            schemaErrors: validation.errors.map((e) => e.stack),
          });
        }
        const video_game = await Video_Game.create(req.body);
        return res.status(201).json({
          message: `${req.body.title} successfully added to database.`,
          saved_video_game: video_game,
        });
      } catch (err) {
        return next(err);
      }
    case "album":
      try {
        const validation = validate(req.body, otherMediaCreateSchema);
        if (!validation.valid) {
          return next({
            status: 400,
            schemaErrors: validation.errors.map((e) => e.stack),
          });
        }
        const album = await Album.create(req.body);
        return res.status(201).json({
          message: `${req.body.title} successfully added to database.`,
          saved_album: album,
        });
      } catch (err) {
        return next(err);
      }
  }
});

router.get("/search", async (req, res, next) => {
  switch (req.query.type) {
    case "book":
      try {
        const bookList = await Book.find(req.query.title);
        if (bookList.length === 0) {
          return next({
            status: 404,
            error: "Media not found",
            message: `No book in database with title ${req.query.title}`,
          });
        }
        return res.status(200).json({
          message: `Successfully found ${bookList.length} book(s) with title ${bookList[0].title}`,
          foundBooksList: bookList,
        });
      } catch (err) {
        return next(err);
      }
    case "movie":
      try {
        const movieList = await Movie.find(req.query.title);
        return res
          .status(200)
          .json({ message: "Successfully ", payload: movieList });
      } catch (err) {
        return next(err);
      }
    case "video_game":
      try {
        const VGList = await Video_Game.find(req.query.title);
        return res
          .status(200)
          .json({ message: "Successful search", payload: VGList });
      } catch (err) {
        return next(err);
      }
    case "album":
      try {
        const albumList = await Album.find(req.query.title);
        return res
          .status(200)
          .json({ message: "Successful search", payload: albumList });
      } catch (err) {
        return next(err);
      }
  }
});

module.exports = router;
