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
                page_count: 100,
                pub_year: 2025,
                spine_color: '#hexcode',
                image_urls: ['123url.com'],
              },
            },
            movie: {
              value: {
                title: 'Movie title edit',
                spine_color: '#hexcode',
                image_urls: ['123url.com'],
              },
            },
            video_game: {
              value: {
                title: 'Video Game title edit',
                spine_color: '#hexcode',
                image_urls: ['123url.com'],
              },
            },
            album: {
              value: {
                title: 'Album title edit',
                spine_color: '#hexcode',
                image_urls: ['123url.com'],
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

                editAttemptItem: {
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
                  message: 'Book title edit successfully edited',
                  saveAttemptItem: {
                    id: 'uuid string',
                    title: 'Book title edit',
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
                  message: 'Movie title edit successfully edited',
                  saveAttemptItem: {
                    id: 'uuid string',
                    title: 'Movie title edit',
                    spine_color: '#hexcode',
                    image_urls: ['123url.com'],
                  },
                  actionCompleted: true,
                  type: 'movie',
                },
              },
              video_game: {
                value: {
                  message: 'Video Game title edit successfully edited',
                  saveAttemptItem: {
                    id: 'uuid string',
                    title: 'Video Game title edit',
                    spine_color: '#hexcode',
                    image_urls: ['123url.com'],
                  },
                  actionCompleted: true,
                  type: 'video_game',
                },
              },
              album: {
                value: {
                  message: 'Album title edited successfully edited',
                  saveAttemptItem: {
                    id: 'uuid string',
                    title: 'Album title edited',
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
        description: 'Item not edited',
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
                editAttemptItem: {
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
                title: 'Album title edit',
                spine_color: '#hexcode',
                image_urls: ['123url.com'],
              },
              actionCompleted: false,
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
                editAttemptItem: {
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
              error: 'Media not found',
              message:
                'Edit requested on an item that does not exist in the database',
              saveAttemptItem: {
                title: 'Album title edit',
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

module.exports = databaseEdit;
