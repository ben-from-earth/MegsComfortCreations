const onlineOpenLibrary = {
  post: {
    tags: ['Get Online Data'],
    summary: 'Gets data from Open Library API',
    description:
      'Using title and author, collect page count and publication year of book from Open Library',

    requestBody: {
      required: true,
      content: {
        'application/json': {
          schema: {
            type: 'object',
            properties: {
              title: { type: 'string' },
              author: { type: 'string' },
            },
            required: ['title', 'author'],
          },
          example: {
            title: 'Dune',
            author: 'Frank Herbert',
          },
        },
      },
    },
    responses: {
      200: {
        description: 'Data successfully gathered',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                title: { type: 'string' },
                author: { type: 'string' },
                pageCount: { type: 'string' },
                pubYear: { type: 'string' },
              },
            },
            example: {
              title: 'Dune',
              author: 'Frank Herbert',
              pageCount: 584,
              pubYear: 1965,
            },
          },
        },
      },
      400: {
        description: 'Open Library error or data not found',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                error: { type: 'string' },
                message: { type: 'string' },
                failedSearchData: {
                  type: 'object',
                  properties: {
                    title: { type: 'string' },
                    author: { type: 'string' },
                  },
                },
              },
            },
            example: {
              error: 'Open Library Error',
              message: `Error gathering Open Library data for [title]`,
              failedSearchData: { title: 'title', author: 'author' },
            },
          },
        },
      },
    },
  },
};

export default onlineOpenLibrary;
