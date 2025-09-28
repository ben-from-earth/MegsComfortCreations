const genresRemoveAllLinks = {
  get: {
    tags: ['Genres'],
    summary: 'Remove all genre links for a book',
    description: 'Using bookID, remove all genre links to the ID',
    parameters: [
      {
        in: 'query',
        name: 'bookID',
        required: true,
        description: 'bookID of book to remove genre links',
        schema: {
          type: 'string',
          enum: ['Purposefully Blank'],
        },
      },
    ],
    responses: {
      200: {
        description: 'Succesful genre link removal',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                message: { type: 'string' },
                actionCompleted: {
                  type: 'boolean',
                },
              },
            },
            example: {
              message: 'All genre links removed for bookID [uuid string]',
              actionCompleted: true,
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

module.exports = genresRemoveAllLinks;
