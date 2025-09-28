const genresPagination = {
  get: {
    tags: ['Genres'],
    summary: 'Paginated media collection',
    description:
      'Using genre, limit, page, sort, asc/desc, collect a list of media rows from the database based on the parameters.',
    parameters: [
      {
        in: 'query',
        name: 'genre',
        required: true,
        description: 'genre of book being searched',
        schema: {
          type: 'string',
          enum: ['Science Fiction', 'Classic Literature'],
        },
      },
      {
        in: 'query',
        name: 'limit',
        required: true,
        description: 'Number of results in the paginated collection.',
        schema: {
          type: 'integer',
          enum: [3, 5, 10],
        },
      },
      {
        in: 'query',
        name: 'page',
        required: true,
        description: 'Page request for pagination',
        schema: {
          type: 'integer',
          enum: [1, 2, 3],
        },
      },
      {
        in: 'query',
        name: 'sort',
        required: true,
        description: 'Which column to sort the database on',
        schema: {
          type: 'string',
          enum: ['title', 'pages'],
        },
      },
      {
        in: 'query',
        name: 'ascDesc',
        required: true,
        description: 'Ascending order or Descending order',
        schema: {
          type: 'string',
          enum: ['asc', 'desc'],
        },
      },
    ],
    responses: {
      200: {
        description: 'List of media rows corresponding to inputted parameters',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                message: { type: 'string' },

                paginatedList: {
                  type: 'array',
                  description:
                    'List of objects from database corresponding to pagination parameters',
                  items: {
                    type: 'object',
                    description: 'Database Item',
                    properties: {
                      id: { type: 'string' },
                      title: { type: 'string' },
                      author: { type: 'string' },
                      page_count: { type: 'integer' },
                      pub_year: { type: 'integer' },
                      spine_color: { type: 'string' },
                      image_urls: {
                        type: 'array',
                        items: {
                          type: 'string',
                          description: 'urls for image of cover',
                        },
                      },
                    },
                  },
                },
                total: {
                  type: 'integer',
                  description:
                    'Total number of items in the corresponding type table',
                },
              },
            },
            example: {
              message: 'Successful database gather',
              paginatedList: [
                {
                  id: 'uuid string',
                  title: 'Book 1',
                  author: 'Author',
                  page_count: 100,
                  pub_year: 2025,
                  spine_color: '#hexcode',
                  image_urls: ['123url.com'],
                },
                {
                  id: 'uuid string',
                  title: 'Book 2',
                  author: 'Author',
                  page_count: 100,
                  pub_year: 2025,
                  spine_color: '#hexcode',
                  image_urls: ['123url.com'],
                },
                {
                  id: 'uuid string',
                  title: 'Book 3',
                  author: 'Author',
                  page_count: 100,
                  pub_year: 2025,
                  spine_color: '#hexcode',
                  image_urls: ['123url.com'],
                },
              ],
              total: 25,
            },
          },
        },
      },
      400: {
        description: 'Database Error',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                error: {
                  type: 'string',
                },
                message: {
                  type: 'string',
                },
              },
            },
            example: {
              error: 'Genre Error',
              message: 'Error connecting to the database and/or genre table',
            },
          },
        },
      },
    },
  },
};

module.exports = genresPagination;
