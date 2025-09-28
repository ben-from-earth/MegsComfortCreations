const express = require('express');
const router = new express.Router();
const outputPNG = require('../helpers/outputPNG');

router.post('/create', async (req, res) => {
  //req body: {template, images: [array of image blocks]}
  //image blocks: {url: "url.com", spine_color: "#ffffffff", type}
  const template = req.body.template;
  const images = req.body.images;
  const png = await outputPNG({ template, images });
  res.setHeader('Content-Type', 'image/png');
  res.status(201).send(png);
});

module.exports = router;
