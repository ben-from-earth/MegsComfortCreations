const databaseGet = {
  get: {
    tags: ['Database'],
    summary: 'Paginated media collection',
    description:
      'Using type, limit, page, sort, asc/desc, collect a list of media rows from the database based on the parameters.',
    parameters: [
      {
        in: 'query',
        name: 'type',
        required: true,
        description: 'Type of media being searched',
        schema: {
          type: 'string',
          enum: ['book', 'movie', 'videoGame', 'album'],
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
          enum: ['title', 'author', 'pageCount', 'pubYear', 'spineColor'],
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
                      pageCount: { type: 'integer' },
                      pubYear: { type: 'integer' },
                      spineColor: { type: 'string' },
                      imageUrls: {
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
                  pageCount: 100,
                  pubYear: 2025,
                  spineColor: '#hexcode',
                  imageUrls: ['123url.com'],
                },
                {
                  id: 'uuid string',
                  title: 'Book 2',
                  author: 'Author',
                  pageCount: 100,
                  pubYear: 2025,
                  spineColor: '#hexcode',
                  imageUrls: ['123url.com'],
                },
                {
                  id: 'uuid string',
                  title: 'Book 3',
                  author: 'Author',
                  pageCount: 100,
                  pubYear: 2025,
                  spineColor: '#hexcode',
                  imageUrls: ['123url.com'],
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
              error: 'Database Collection Error',
              message:
                'Error Gathering items from the database during pagination',
            },
          },
        },
      },
    },
  },
};

export default databaseGet;
