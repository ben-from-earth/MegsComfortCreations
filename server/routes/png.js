const express = require('express');
const router = new express.Router();
const { outputAuto } = require('../helpers/outputPNG');

router.post('/create', async (req, res) => {
  try {
    //req body: {template, images: [array of image blocks]}
    //image blocks: {url: "url.com", spine_color: "#ffffffff", type}
    const { template, images } = req.body;
    const { mime, filename, buffer } = await outputAuto({
      template,
      images,
      prefix: 'grid',
    });
    res.setHeader('Content-Type', mime);
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.status(201).send(buffer);
  } catch (error) {
    next(error);
  }
});

module.exports = router;
