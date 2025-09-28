const genresUnlink = {
  post: {
    tags: ['Genres'],
    summary: 'Remove a book-genre link',
    description:
      'Provided a bookID and list of genres, remove links from genres_books link table',
    requestBody: {
      required: true,
      content: {
        'application/json': {
          schema: {
            type: 'object',
            properties: {
              bookID: {
                type: 'string',
                description: 'uuid string of book to remove link to genres',
              },
              genres: {
                type: 'array',
                description: 'list of genres to remove link to bookID',
                items: {
                  type: 'string',
                },
              },
            },
            required: ['bookID', 'genres'],
          },
          example: {
            bookID: 'uuid string',
            genres: ['Science Fiction', 'Fantasy'],
          },
        },
      },
    },
    responses: {
      200: {
        description: 'bookID and genres successfully unlinked',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                responses: {
                  type: 'array',
                  description: 'List of responses related to genre unlink',
                  items: {
                    type: 'object',
                    properties: {
                      message: { type: 'string' },
                    },
                  },
                },
              },
            },
            example: {
              responses: [
                {
                  message: 'Successfully removed genre: Science Fiction',
                },
                {
                  message: 'Successfully removed genre: Fantasy',
                },
              ],
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

module.exports = genresUnlink;
