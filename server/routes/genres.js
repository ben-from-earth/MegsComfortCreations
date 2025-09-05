const express = require("express");
const router = new express.Router();

const Genre = require("../models/genre");

router.get("/getAll", async (req, res, next) => {
  try {
    const genres = await Genre.getAllGenres();
    return res.status(200).json({ message: "success", payload: genres });
  } catch (err) {
    return next(err);
  }
});

router.post("/getFromBook", async (req, res, next) => {
  try {
    const bookID = req.body.bookID;
    const genres = await Genre.getFromBook(bookID);
    return res
      .status(200)
      .json({ message: "Successfully grabbed genres", payload: genres });
  } catch (err) {
    return next(err);
  }
});

router.post("/addlink", async (req, res, next) => {
  const bookID = req.body.bookID;
  const genres = req.body.genres;
  let response = [];
  for (let genre of genres) {
    try {
      await Genre.link(genre, bookID);
      response.push({ message: "Successful genre link", genre, bookID });
    } catch (err) {
      return next(err);
    }
  }
  return res.status(201).json({ responses: response });
});

module.exports = router;
