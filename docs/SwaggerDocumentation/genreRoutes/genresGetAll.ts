const genresGetAll = {
  get: {
    tags: ['Genres'],
    summary: 'Get all genres from the genre table',
    description: 'Get all genres from the genre table',
    responses: {
      200: {
        description: 'List of all genres',
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
              message: 'Success',
              genres: [
                "Children's Fiction",
                'Middle-Grade Fiction',
                'Young Adult Fiction',
                'New Adult Fiction',
                'Romance',
                'Contemporary Fiction',
                'Spicy Romance (18+)',
                'LGBTQ',
                'Romantasy',
                'Fantasy',
                'Historical Fiction',
                'Mystery',
                'Thriller',
                'Horror',
                'Science Fiction',
                'Classic Literature',
                'Memoir',
                'History',
                'Philosophy',
                'Anthology',
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

export default genresGetAll;
