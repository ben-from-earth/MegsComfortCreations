// swagger.js
const swaggerJsdoc = require('swagger-jsdoc');

//schemas
const bookSaveSchema = require('../schemas/bookCreateSchema.json');
const otherSaveSchema = require('../schemas/otherMediaCreateSchema.json');

//database routes
const databaseSave = require('./databaseRoutes/databaseSave');
const databaseSearch = require('./databaseRoutes/databaseSearch');
const databaseGet = require('./databaseRoutes/databaseGet');
const databaseEdit = require('./databaseRoutes/databaseEdit');
const databaseDelete = require('./databaseRoutes/databaseDelete');

//genre routes
const genresGetAll = require('./genreRoutes/genresGetAll');
const genresGetForBook = require('./genreRoutes/genresGetForBook');
const genresAddLink = require('./genreRoutes/genresAddLink');
const genresUnlink = require('./genreRoutes/genresUnlink');
const genresPagination = require('./genreRoutes/genresPagination');
const genresNoGenres = require('./genreRoutes/genresNoGenres');
const genresRemoveAllLinks = require('./genreRoutes/genresRemoveAllLinks');

//png routes
const pngCreate = require('./pngRoutes/pngCreate');

//online data collection routes
const onlineOpenLibrary = require('./getOnlineDataRoutes/onlineOpenLibrary');
const onlineMediaCovers = require('./getOnlineDataRoutes/onlineMediaCovers');

const definition = {
  openapi: '3.0.3',
  info: {
    title: 'Megs Comfort Creations API',
    version: '1.0.0',
    description: 'API for Media Collector',
  },
  servers: [{ url: 'http://localhost:3001', description: 'Local dev' }],
  tags: [
    {
      name: 'Database',
      description:
        'Creating, editing, and searching items in the related database',
    },
    {
      name: 'Genres',
      description:
        'Dealing with genres (get all and adding/removing book-genre links)',
    },
    {
      name: 'Get Online Data',
      description:
        'Routes that interact with publically available APIs: Google Search and OpenLibrary',
    },
    {
      name: 'PNG Creation',
      description:
        'Creating the final PNG file from the collected media cover images',
    },
  ],
  components: {
    schemas: {
      Book: bookSaveSchema,
      OtherMedia: otherSaveSchema,
    },
  },
  paths: {
    '/database/save/{type}': databaseSave,
    '/database/search': databaseSearch,
    '/database': databaseGet,
    '/database': databaseDelete,
    '/database/edit/{type}': databaseEdit,
    '/genres/getAll': genresGetAll,
    '/genres/getForBook': genresGetForBook,
    '/genres/addLink': genresAddLink,
    '/genres/unlink': genresUnlink,
    '/genres': genresPagination,
    '/genres/noGenres': genresNoGenres,
    '/genres/removeAllLinksForBook': genresRemoveAllLinks,
    '/png/create': pngCreate,
    '/getOnlineData/openlibrary': onlineOpenLibrary,
    '/getOnlineData/mediacovers': onlineMediaCovers,
  },
};

// Tell swagger-jsdoc where to find your JSDoc annotations
const options = {
  definition,
  apis: ['./routes/**/*.js', './app.js'], // adjust to your folders
};

const openapiSpec = swaggerJsdoc(options);

module.exports = { openapiSpec };
