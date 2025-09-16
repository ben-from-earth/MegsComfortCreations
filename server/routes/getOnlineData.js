const express = require("express");
const router = new express.Router();
const axios = require("axios");

const API_KEY = process.env.GOOGLE_SEARCH_API_KEY;
const CX = process.env.GOOGLE_SEARCH_CX;

router.post("/openlibrary", async (req, res, next) => {
  const title = req.body.title;
  const author = req.body.author;

  try {
    const params = new URLSearchParams({
      title,
      author,
      limit: "1",
      fields: "first_publish_year,number_of_pages_median",
    });
    const openLibraryRes = await axios.get(
      `https://openlibrary.org/search.json?${params.toString()}`
    );

    const data = openLibraryRes.data;
    const doc = data?.docs?.[0];
    if (!doc) {
      openLibraryRes.json({ title, author });
    }

    const { first_publish_year: pub_year, number_of_pages_median: page_count } =
      doc;
    res.status(200).json({ title, author, pub_year, page_count });
  } catch {
    next({
      status: 400,
      error: "Open Library Error",
      message: `Error gathering Open Library data for ${title}`,
      failedSearchData: { title, author },
    });
  }
});

router.post("/mediacovers", async (req, res, next) => {
  const title = req.body.title;
  const type = req.body.type;
  const author = req.body.author;

  const imgArr = [];

  if (!CX || !API_KEY) {
    next({
      status: 400,
      error: "Google Search Credential Error",
      message:
        "Error connecting to Google Search API because of invalid or empty credentials",
    });
  }

  try {
    const params = new URLSearchParams({
      q: `${title}${author ? ` ${author}` : ""} ${type} Cover Image`,
      cx: CX,
      key: API_KEY,
      searchType: "image",
      num: 3,
    });

    const imageSearchRes = await axios.get(
      `https://www.googleapis.com/customsearch/v1?${params.toString()}`
    );

    const imageURLs = imageSearchRes.data;
    imageURLs.items.map((i) => imgArr.push(i.link));
    res.status(200).json({ images: imgArr });
  } catch {
    next({
      status: 400,
      error: "Google Search Error",
      message: "Error connecting to Google Search API",
    });
  }
});

module.exports = router;
