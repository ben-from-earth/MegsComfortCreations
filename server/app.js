const express = require("express");
const cors = require("cors");

const outputPNG = require("./outputPNG");

const ExpressError = require("./expressError");

//get routes
const databaseRoutes = require("./routes/database");
const genresRoutes = require("./routes/genres");

const app = express();

app.use(
  cors({
    origin: "http://localhost:5173",
    methods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);

app.use(express.json());
app.use("/database", databaseRoutes);
app.use("/genres", genresRoutes);

app.post("/print-png", async (req, res) => {
  //req body: {template, images: [array of image blocks]}
  //image blocks: {url: "url.com", spine_color: "#ffffffff", type}
  const template = req.body.template;
  const images = req.body.images;
  const png = await outputPNG({ template, images });
  res.setHeader("Content-Type", "image/png");
  res.send(png);
});

app.use((req, res, next) => {
  const err = new ExpressError("Not Found", 404);
  return next(err);
});

app.use((err, req, res, next) => {
  res.status(err.status || 500);
  let errorResponse = { error: "", message: "" };

  if (err.error === "Media not found") {
    errorResponse.error = err.error;
    errorResponse.message = err.message;
  }
  //errors from schema violation
  if (err.schemaErrors) {
    const missingFields = [];
    const wrongTypes = [];
    const arrayLengthViolation = [];
    for (let error of err.schemaErrors) {
      if (error.includes("instance requires property")) {
        const missingField = error.split('"')[1];
        missingFields.push(
          `Database save request is missing field: ${missingField}`
        );
      } else if (error.includes("is not of a type(s)")) {
        const wrongTypeField = error.split(" ")[0].split(".")[1];
        const correctType = error.split(") ")[1];
        wrongTypes.push(
          `${wrongTypeField} was input as the wrong type, should be a(n) ${correctType}`
        );
      } else if (error.includes("does not meet minimum length of 1")) {
        arrayLengthViolation.push("Database save attempted without images");
      }
    }
    errorResponse.error = "Schema violation during save request";
    errorResponse["validationErrors"] = [
      ...missingFields,
      ...wrongTypes,
      ...arrayLengthViolation,
    ];
  }

  //errors from pg
  if (err.detail) {
    let errorDetail = err.detail;
    if (errorDetail?.includes("Failing row")) console.log(err);

    if (
      errorDetail?.includes("Key (title, author)") &&
      errorDetail?.includes("already exists")
    ) {
      errorResponse.error = "Error saving book to database";
      res.status(400);
    }

    errorResponse.message = err.detail;
  }

  return res.json(errorResponse);
});

module.exports = app;
