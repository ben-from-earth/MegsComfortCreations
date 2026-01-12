// swagger.js
import swaggerJsdoc from 'swagger-jsdoc';

//schemas
import bookSaveSchema from 'lib/database/schemas/bookCreateSchema.json';
import otherSaveSchema from 'lib/database/schemas/otherMediaCreateSchema.json';

//database routes
import databaseSave from '@databaseDocs/databaseSave';
import databaseSearch from '@databaseDocs/databaseSearch';
import databaseGet from '@databaseDocs/databaseGet';
import databaseEdit from '@databaseDocs/databaseEdit';
import databaseDelete from '@databaseDocs/databaseDelete';

//genre routes
import genresGetAll from '@genreDocs/genresGetAll';
import genresGetForBook from '@genreDocs/genresGetForBook';
import genresAddLink from '@genreDocs/genresAddLink';
import genresUnlink from '@genreDocs/genresUnlink';
import genresPagination from '@genreDocs/genresPagination';
import genresNoGenres from '@genreDocs/genresNoGenres';

//png routes
import pngCreate from '@pngDocs/pngCreate';

//online data collection routes
import onlineOpenLibrary from '@onlineDataDocs/onlineOpenLibrary';
import onlineMediaCovers from '@onlineDataDocs/onlineMediaCovers';
import path from 'path';
import { NextResponse } from 'next/server';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

const definition = {
  openapi: '3.0.3',
  info: {
    title: 'Megs Comfort Creations API',
    version: '1.0.0',
    description: 'API for Media Collector',
  },
  servers: [{ url: 'http://localhost:3000', description: 'Local dev' }],
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
    '/api/database/save/{type}': databaseSave,
    '/api/database/search': databaseSearch,
    '/api/database': databaseGet,
    '/api/database/delete': databaseDelete,
    '/api/database/edit/{type}': databaseEdit,
    '/api/genres/getall': genresGetAll,
    '/api/genres/getforbook': genresGetForBook,
    '/api/genres/addlink': genresAddLink,
    '/api/genres/unlink': genresUnlink,
    '/api/genres': genresPagination,
    '/api/genres/nogenres': genresNoGenres,
    '/api/png/create': pngCreate,
    '/api/getonlinedata/openlibrary': onlineOpenLibrary,
    '/api/getonlinedata/mediacovers': onlineMediaCovers,
  },
};

// Tell swagger-jsdoc where to find your JSDoc annotations
const options = {
  definition,
  apis: [path.join(process.cwd(), 'app/api/**/*.ts')], // adjust to your folders
};

const openapiSpec = swaggerJsdoc(options);

export function GET() {
  return NextResponse.json(openapiSpec);
}
