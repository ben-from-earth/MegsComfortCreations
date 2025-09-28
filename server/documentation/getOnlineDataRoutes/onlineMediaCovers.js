const onlineMediaCovers = {
  post: {
    tags: ['Get Online Data'],
    summary: 'Gets data from Google Search API',
    description:
      'Using title, type, and (optionally) author to collect image URLs from Google Search.',
    requestBody: {
      required: true,
      content: {
        'application/json': {
          schema: {
            type: 'object',
            properties: {
              title: { type: 'string' },
              author: { type: 'string' },
              type: {
                type: 'string',
                enum: ['book', 'movie', 'video_game', 'album'],
              },
            },
            required: ['title', 'type'],
          },
          example: { title: 'Dune', author: 'Frank Herbert', type: 'book' },
        },
      },
    },
    responses: {
      200: {
        description: 'Data successfully gathered',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                images: {
                  type: 'array',
                  items: { type: 'string', format: 'uri' },
                },
              },
              required: ['images'],
            },
            example: {
              images: ['url1.com', 'url2.com', 'url3.com'],
            },
          },
        },
      },
      400: {
        description: 'Bad Request',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                error: { type: 'string' },
                message: { type: 'string' },
              },
              required: ['error', 'message'],
            },
            examples: {
              credentialError: {
                summary: 'Missing/invalid Google credentials',
                value: {
                  error: 'Google Search Credential Error',
                  message:
                    'Error connecting to Google Search API because of invalid or empty credentials',
                },
              },
              apiError: {
                summary: 'Upstream API error',
                value: {
                  error: 'Google Search Error',
                  message: 'Error connecting to Google Search API',
                },
              },
            },
          },
        },
      },
    },
  },
};

module.exports = onlineMediaCovers;
