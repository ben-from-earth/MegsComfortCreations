const databaseEdit = {
  put: {
    tags: ['Database'],
    summary: 'Edits a media item in the database',
    description:
      'Given a request body, overwrites the information in the database for that row',
    parameters: [
      {
        in: 'path',
        name: 'type',
        required: true,
        description: 'Type of media being edited',
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
                title: 'Book title edit',
                author: 'Book author',
                pageCount: 100,
                pubYear: 2025,
                spineColor: '#hexcode',
                imageUrls: ['123url.com'],
              },
            },
            movie: {
              value: {
                title: 'Movie title edit',
                spineColor: '#hexcode',
                imageUrls: ['123url.com'],
              },
            },
            video_game: {
              value: {
                title: 'Video Game title edit',
                spineColor: '#hexcode',
                imageUrls: ['123url.com'],
              },
            },
            album: {
              value: {
                title: 'Album title edit',
                spineColor: '#hexcode',
                imageUrls: ['123url.com'],
              },
            },
          },
        },
      },
    },
    responses: {
      200: {
        description: 'Item successfully edited.',
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
                  description: 'Type of media that was edited.',
                },
              },
            },
            examples: {
              book: {
                value: {
                  message: 'Book title edit successfully edited',
                  actionAttemptItem: {
                    id: 'uuid string',
                    title: 'Book title edit',
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
                  message: 'Movie title edit successfully edited',
                  actionAttemptItem: {
                    id: 'uuid string',
                    title: 'Movie title edit',
                    spineColor: '#hexcode',
                    imageUrls: ['123url.com'],
                  },
                  type: 'movie',
                },
              },
              video_game: {
                value: {
                  message: 'Video Game title edit successfully edited',
                  actionAttemptItem: {
                    id: 'uuid string',
                    title: 'Video Game title edit',
                    spineColor: '#hexcode',
                    imageUrls: ['123url.com'],
                  },
                  type: 'video_game',
                },
              },
              album: {
                value: {
                  message: 'Album title edited successfully edited',
                  actionAttemptItem: {
                    id: 'uuid string',
                    title: 'Album title edited',
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
        description: 'Item not edited',
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
                  description: 'Type of media that was edited.',
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
      404: {
        description: 'Item requested to edit doesnt exist in the database',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                error: { type: 'string' },
                message: {
                  type: 'string',
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
                  description: 'Type of media that was edited.',
                },
                errors: {
                  type: 'array',
                  items: {
                    type: 'string',
                  },
                },
              },
            },
            example: {
              error: 'Media not found',
              message:
                'Edit requested on an item that does not exist in the database',
              actionAttemptItem: {
                title: 'Album title edit',
                spineColor: '#hexcode',
                imageUrls: ['123url.com'],
              },
              type: 'album',
              errors: ['[title] does not exist in the database'],
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
                error: { type: 'string' },
                message: {
                  type: 'string',
                },
              },
            },
            example: {
              error: 'Edit Error',
              message: 'Edit request of [title] failed',
            },
          },
        },
      },
    },
  },
};

export default databaseEdit;
