const databaseSave = {
  post: {
    tags: ['Database'],
    summary: 'Saves a media item to the database',
    description: 'Creates a row for the specified media type.',
    parameters: [
      {
        in: 'path',
        name: 'type',
        required: true,
        description: 'Type of media being saved',
        schema: {
          type: 'string',
          enum: ['book', 'movie', 'video_game', 'album'],
        },
      },
    ],
    requestBody: {
      required: true,
      content: {
        'application/json': {
          schema: {
            oneOf: [
              { $ref: '../components/schemas/Book' },
              { $ref: '../components/schemas/OtherMedia' },
            ],
          },
          examples: {
            book: {
              value: {
                title: 'Book title',
                author: 'Book author',
                page_count: 100,
                pub_year: 2025,
                spine_color: '#hexcode',
                image_urls: ['123url.com'],
              },
            },
            movie: {
              value: {
                title: 'Movie title',
                spine_color: '#hexcode',
                image_urls: ['123url.com'],
              },
            },
            video_game: {
              value: {
                title: 'Video Game title',
                spine_color: '#hexcode',
                image_urls: ['123url.com'],
              },
            },
            album: {
              value: {
                title: 'Album title',
                spine_color: '#hexcode',
                image_urls: ['123url.com'],
              },
            },
          },
        },
      },
    },
    responses: {
      201: {
        description: 'Item successfully saved.',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                message: { type: 'string' },

                saveAttemptItem: {
                  type: 'object',
                  description:
                    'Request body returned with id generated from the database.',
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
                actionCompleted: {
                  type: 'boolean',
                  description: 'true because action (save) was completed',
                },
                type: {
                  type: 'string',
                  description: 'Type of media that was saved.',
                },
              },
            },
            examples: {
              book: {
                value: {
                  message: 'Book title successfully added to the database',
                  saveAttemptItem: {
                    id: 'uuid string',
                    title: 'Book title',
                    author: 'Book author',
                    page_count: 100,
                    pub_year: 2025,
                    spine_color: '#hexcode',
                    image_urls: ['123url.com'],
                  },
                  actionCompleted: true,
                  type: 'book',
                },
              },
              movie: {
                value: {
                  message: 'Movie title successfully added to the database',
                  saveAttemptItem: {
                    id: 'uuid string',
                    title: 'Movie title',
                    spine_color: '#hexcode',
                    image_urls: ['123url.com'],
                  },
                  actionCompleted: true,
                  type: 'movie',
                },
              },
              video_game: {
                value: {
                  message:
                    'Video Game title successfully added to the database',
                  saveAttemptItem: {
                    id: 'uuid string',
                    title: 'Video Game title',
                    spine_color: '#hexcode',
                    image_urls: ['123url.com'],
                  },
                  actionCompleted: true,
                  type: 'video_game',
                },
              },
              album: {
                value: {
                  message: 'Album title successfully added to the database',
                  saveAttemptItem: {
                    id: 'uuid string',
                    title: 'Album title',
                    spine_color: '#hexcode',
                    image_urls: ['123url.com'],
                  },
                  actionCompleted: true,
                  type: 'album',
                },
              },
            },
          },
        },
      },
      400: {
        description: 'Item not saved',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                schemaErrors: {
                  type: 'array',
                  description: 'List of schema errors',
                  items: {
                    type: 'string',
                    description: 'Schema error',
                  },
                },
                saveAttemptItem: {
                  type: 'object',
                  description: 'Request body returned',
                  properties: {
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
                actionCompleted: {
                  type: 'boolean',
                  description: 'false because action (save) was not completed',
                },
                type: {
                  type: 'string',
                  description: 'Type of media that was saved.',
                },
              },
            },
            example: {
              schemaErrors: [
                'Save/edit attempt missing spine_color',
                'Save/edit attempt missing image_urls',
              ],
              saveAttemptItem: {
                title: 'Album title',
                spine_color: '#hexcode',
                image_urls: ['123url.com'],
              },
              actionCompleted: false,
              type: 'album',
            },
          },
        },
      },
    },
  },
};

module.exports = databaseSave;
