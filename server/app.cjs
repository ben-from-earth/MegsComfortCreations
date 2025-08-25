const express = require("express");
const cors = require("cors");

//Get all models
const Book = require("./models/book.cjs");
const Movie = require("./models/movie.cjs");
const Video_Game = require("./models/video_game.cjs");
const Album = require("./models/album.cjs");
const Genre = require("./models/genre.cjs");
const outputPNG = require("./outputPNG.cjs");

const app = express();

app.use(
  cors({
    origin: "http://localhost:5173",
    methods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);

app.use(express.json());

const ExpressError = require("./expressError.cjs");

app.get("/genres", async (req, res, next) => {
  try {
    const genres = await Genre.getAllGenres();
    return res.status(200).json({ outcome: "success", payload: genres });
  } catch (err) {
    return next(err);
  }
});

app.post("/print-png", async (req, res) => {
  //req body: {template, images: [array of image blocks]}
  //image blocks: {url: "url.com", spineColor: "#ffffffff", type}
  const template = req.body.template;
  const images = req.body.images;
  const png = await outputPNG({ template, images });
  res.setHeader("Content-Type", "image/png");
  res.send(png);
});

app.post("/savetodb/:type", async (req, res, next) => {
  switch (req.params.type) {
    case "book":
      try {
        const book = await Book.create(req.body);
        //book.id will give in database id
        return res.status(201).json({ created: "success", payload: book });
      } catch (err) {
        return next(err);
      }
    case "movie":
      try {
        const movie = await Movie.create(req.body);
        return res.status(201).json({ created: "success", payload: movie });
      } catch (err) {
        return next(err);
      }
    case "video_game":
      try {
        const video_game = await Video_Game.create(req.body);
        return res
          .status(201)
          .json({ created: "success", payload: video_game });
      } catch (err) {
        return next(err);
      }
    case "album":
      try {
        const album = await Album.create(req.body);
        return res.status(201).json({ created: "success", payload: album });
      } catch (err) {
        return next(err);
      }
  }
});

app.use((req, res, next) => {
  const err = new ExpressError("Not Found", 404);
  return next(err);
});

app.use((err, req, res, next) => {
  res.status(err.status || 500);

  return res.json({
    error: err,
    message: err.message,
  });
});

module.exports = app;
