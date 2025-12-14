const databaseSearch = {
  get: {
    tags: ['Database'],
    summary: 'Find a media item in the database',
    description:
      'Using type and title, search the corresponding type table for the row with corresponding title',
    parameters: [
      {
        in: 'query',
        name: 'type',
        required: true,
        description: 'Type of media being searched',
        schema: {
          type: 'string',
          enum: ['book', 'movie', 'video_game', 'album'],
        },
      },
      {
        in: 'query',
        name: 'title',
        required: true,
        description: 'Title of media being searched',
        schema: {
          type: 'string',
          enum: [
            'Dune',
            'Avatar',
            'Rocket League',
            'The Dark Side of the Moon',
          ],
        },
      },
    ],
    responses: {
      200: {
        description: 'Found item in the database with type/title combo',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                message: { type: 'string' },

                foundMediaList: {
                  type: 'array',
                  description:
                    'List of objects from database corresponding to title/type combo (could be multiple)',
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
                    'Total number of items in the database fitting the title/type combo',
                },
              },
            },
            examples: {
              book: {
                value: {
                  message: 'Successfully found 1 book(s) with title Dune',
                  foundMediaList: [
                    {
                      id: 'uuid string',
                      title: 'Dune',
                      author: 'Frank Herbert',
                      pageCount: 592,
                      pubYear: 1965,
                      spineColor: '#f25b26',
                      imageUrls: ['123url.com'],
                    },
                  ],
                  total: 1,
                },
              },
              movie: {
                value: {
                  message: 'Successfully found 1 movie(s) with title Avatar',
                  foundMediaList: [
                    {
                      id: 'uuid string',
                      title: 'Avatar',
                      spineColor: '#000000',
                      imageUrls: ['123url.com'],
                    },
                  ],
                  total: 1,
                },
              },
              video_game: {
                value: {
                  message:
                    'Successfully found 1 video_game(s) with title Rocket League',
                  foundMediaList: [
                    {
                      id: 'uuid string',
                      title: 'Rocket League',
                      spineColor: '#587cba',
                      imageUrls: ['123url.com'],
                    },
                  ],
                  total: 1,
                },
              },
              album: {
                value: {
                  message:
                    'Successfully found 1 album(s) with title The Dark Side of the Moon',
                  foundMediaList: [
                    {
                      id: 'uuid string',
                      title: 'Dark Side of the Moon, The',
                      imageUrls: ['123url.com'],
                    },
                  ],
                  total: 1,
                },
              },
            },
          },
        },
      },
      404: {
        description: 'Item not found',
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
              error: 'Media not found',
              message: 'No [media] in database with title [Title]',
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
              error: 'Database collection error',
              message: 'Error gathering items from the database during search',
            },
          },
        },
      },
      422: {
        description: 'Malformed Search Query',
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
              error: 'Missing Search Parameters',
              message: 'Search parameters "type" and "title" are required',
            },
          },
        },
      },
    },
  },
};

export default databaseSearch;
