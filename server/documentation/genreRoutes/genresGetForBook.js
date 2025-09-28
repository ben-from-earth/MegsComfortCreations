const genresGetForBook = {
  get: {
    tags: ['Genres'],
    summary: 'Get all the genre linked to a book',
    description:
      'Using bookID get all genres in the table listed for that bookID',
    parameters: [
      {
        in: 'query',
        name: 'bookID',
        required: true,
        description: 'bookID of book looking for genre links (test = Dune)',
        schema: {
          type: 'string',
          enum: ['b2ec0541-1a6f-4789-9c98-e351a9a02784'],
        },
      },
    ],
    responses: {
      200: {
        description: 'List of all genres tied to bookID',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                message: { type: 'string' },

                genres: {
                  type: 'array',
                  description: 'List of genres',
                  items: {
                    type: 'string',
                  },
                },
              },
            },
            example: {
              message: 'Successfully grabbed genres for bookID [uuid string]',
              genres: ['Historical Fiction', 'Fantasy'],
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

module.exports = genresGetForBook;
