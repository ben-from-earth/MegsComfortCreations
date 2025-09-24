const express = require('express');
const cors = require('cors');
const outputPNG = require('./helpers/outputPNG');

//get routes
const databaseRoutes = require('./routes/database');
const genresRoutes = require('./routes/genres');
const getOnlineDataRoutes = require('./routes/getOnlineData');

const app = express();

app.use(
  cors({
    origin: 'http://localhost:5173',
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
  })
);

app.use(express.json());
app.use('/database', databaseRoutes);
app.use('/genres', genresRoutes);
app.use('/getOnlineData', getOnlineDataRoutes);

app.post('/print-png', async (req, res) => {
  //req body: {template, images: [array of image blocks]}
  //image blocks: {url: "url.com", spine_color: "#ffffffff", type}
  const template = req.body.template;
  const images = req.body.images;
  const png = await outputPNG({ template, images });
  res.setHeader('Content-Type', 'image/png');
  res.send(png);
});

app.use((req, res, next) => {
  res.status(404).json({
    error: 'Page not Found',
    message: 'The requested route does not exist',
  });
});

app.use((err, req, res, next) => {
  res.status(err.status || 500);
  let errorResponse = { errors: [], message: '', actionCompleted: false };

  if (err.error === 'Media not found') {
    errorResponse.errors.push(err.error);
    errorResponse.message = err.message;
  } else if (err.error === 'Open Library Error') {
    errorResponse.errors.push(err.error);
    errorResponse.message = err.message;
    errorResponse.failedSearchData = err.failedSearchData;
  } else if (
    //errors from schema violation
    err.schemaErrors
  ) {
    const missingFields = [];
    const wrongTypes = [];
    for (let error of err.schemaErrors) {
      if (error.includes('instance requires property')) {
        const missingField = error.split('"')[1];
        missingFields.push(`Save/Edit request missing ${missingField}`);
      } else if (error.includes('is not of a type(s)')) {
        const wrongTypeField = error.split(' ')[0].split('.')[1];
        wrongTypes.push(`${wrongTypeField} is of wrong type`);
      } else if (error.includes('does not meet minimum length')) {
        const field = error.split(' ')[0].split('.')[1];
        missingFields.push(`Save/Edit request missing ${field}`);
      }
    }
    errorResponse.message = 'Schema violation(s) during save/edit request';
    errorResponse.saveAttemptItem = err.saveAttemptItem;

    errorResponse.errors = [...missingFields, ...wrongTypes];
  } else if (err.detail) {
    //errors from PostgreSQL
    let errorDetail = err.detail;
    if (errorDetail?.includes('Failing row'))
      console.log('Failing Row Error:', err);

    if (errorDetail?.includes('already exists')) {
      errorResponse.errors.push('This media already exists in database');
      res.status(400);
    }

    errorResponse.message = errorDetail;
    errorResponse.saveAttemptItem = err.saveAttemptItem;
  } else {
    errorResponse.errors.push(err.error);
    errorResponse.message = err.message;
  }

  return res.json(errorResponse);
});

module.exports = app;
