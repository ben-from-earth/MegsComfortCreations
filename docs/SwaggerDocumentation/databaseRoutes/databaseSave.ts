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
                pageCount: 100,
                pubYear: 2025,
                spineColor: '#hexcode',
                imageUrls: ['123url.com'],
              },
            },
            movie: {
              value: {
                title: 'Movie title',
                spineColor: '#hexcode',
                imageUrls: ['123url.com'],
              },
            },
            video_game: {
              value: {
                title: 'Video Game title',
                spineColor: '#hexcode',
                imageUrls: ['123url.com'],
              },
            },
            album: {
              value: {
                title: 'Album title',
                spineColor: '#hexcode',
                imageUrls: ['123url.com'],
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

                actionAttemptItem: {
                  type: 'object',
                  description:
                    'Request body returned with id generated from the database.',
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
                  actionAttemptItem: {
                    id: 'uuid string',
                    title: 'Book title',
                    author: 'Book author',
                    pageCount: 100,
                    pubYear: 2025,
                    spineColor: '#hexcode',
                    imageUrls: ['123url.com'],
                  },
                  type: 'book',
                },
              },
              movie: {
                value: {
                  message: 'Movie title successfully added to the database',
                  actionAttemptItem: {
                    id: 'uuid string',
                    title: 'Movie title',
                    spineColor: '#hexcode',
                    imageUrls: ['123url.com'],
                  },
                  type: 'movie',
                },
              },
              video_game: {
                value: {
                  message:
                    'Video Game title successfully added to the database',
                  actionAttemptItem: {
                    id: 'uuid string',
                    title: 'Video Game title',
                    spineColor: '#hexcode',
                    imageUrls: ['123url.com'],
                  },
                  type: 'video_game',
                },
              },
              album: {
                value: {
                  message: 'Album title successfully added to the database',
                  actionAttemptItem: {
                    id: 'uuid string',
                    title: 'Album title',
                    spineColor: '#hexcode',
                    imageUrls: ['123url.com'],
                  },
                  type: 'album',
                },
              },
            },
          },
        },
      },
      422: {
        description: 'Item not saved',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                error: { type: 'string' },
                message: { type: 'string' },
                errors: {
                  type: 'array',
                  description: 'List of schema errors',
                  items: {
                    type: 'string',
                    description: 'Schema error',
                  },
                },
                actionAttemptItem: {
                  type: 'object',
                  description: 'Request body returned',
                  properties: {
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
                type: {
                  type: 'string',
                  description: 'Type of media that was saved.',
                },
              },
            },
            example: {
              error: 'Schema Violation Error',
              message: 'Schema violation(s) during save/edit request',
              errors: [
                'Save/edit attempt missing spineColor',
                'Save/edit attempt missing imageUrls',
              ],
              actionAttemptItem: {
                title: 'Album title',
                spineColor: '#hexcode',
                imageUrls: ['123url.com'],
              },
              type: 'album',
            },
          },
        },
      },
      409: {
        description: 'Duplication Error',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                error: {
                  type: 'string',
                },
                message: { type: 'string' },
                errors: {
                  type: 'array',
                  description: 'List of errors',
                  items: {
                    type: 'string',
                    description: 'Duplication error',
                  },
                },
                actionAttemptItem: {
                  type: 'object',
                  description: 'Request body returned',
                  properties: {
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
                type: {
                  type: 'string',
                  description: 'Type of media that was saved.',
                },
              },
            },
            example: {
              error: 'Duplication Attempt Error',
              message:
                'You attempted to save an item to the database that already exists',
              errors: ['key (title) = [title] already exits'],
              actionAttemptItem: {
                title: 'Album title',
                spineColor: '#hexcode',
                imageUrls: ['123url.com'],
              },
              type: 'album',
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
                message: { type: 'string' },
              },
            },
            example: {
              error: 'Database Error',
              message: 'Database Error during save attempt',
            },
          },
        },
      },
    },
  },
};

export default databaseSave;
