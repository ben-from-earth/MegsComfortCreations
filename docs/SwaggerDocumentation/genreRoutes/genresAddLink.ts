const genresAddLink = {
  post: {
    tags: ['Genres'],
    summary: 'Create a book-genre link',
    description:
      'Provided a bookID and list of genres, add to genres_books link table',
    requestBody: {
      required: true,
      content: {
        'application/json': {
          schema: {
            type: 'object',
            properties: {
              bookID: {
                type: 'string',
                description: 'uuid string of book to link to genres',
              },
              genres: {
                type: 'array',
                description: 'list of genres to link to bookID',
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
      201: {
        description: 'bookID and genres successfully linked',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                genreResponses: {
                  type: 'array',
                  description: 'List of responses related to genre link',
                  items: {
                    type: 'object',
                    properties: {
                      message: { type: 'string' },
                      genre: { type: 'string' },
                      bookID: { type: 'string' },
                    },
                  },
                },
              },
            },
            example: {
              responses: [
                {
                  message: 'Successful genre link',
                  genre: ['Science Fiction'],
                  bookID: 'uuid string',
                },
                {
                  message: 'Successful genre link',
                  genre: ['Fantasy'],
                  bookID: 'uuid string',
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

export default genresAddLink;
