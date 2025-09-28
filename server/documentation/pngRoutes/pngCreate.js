const pngCreate = {
  post: {
    tags: ['PNG Creation'],
    summary: 'Generate a PNG grid from images',
    description:
      'Takes a template number and an array of image blocks, then returns a PNG image.',
    requestBody: {
      required: true,
      content: {
        'application/json': {
          schema: {
            type: 'object',
            properties: {
              template: {
                type: 'integer',
                description: 'Template style ID (3 || 5)',
              },
              images: {
                type: 'array',
                description: 'Array of image blocks to render',
                items: {
                  type: 'object',
                  required: ['url', 'spine_color', 'type'],
                  properties: {
                    url: {
                      type: 'string',
                      format: 'uri',
                      description: 'Source image url',
                    },
                    spine_color: {
                      type: 'string',
                      description: 'spine color hex code (e.g. #ffffff)',
                    },
                    type: {
                      type: 'string',
                      description: 'type of media',
                    },
                  },
                },
              },
            },
            required: ['template', 'images'],
          },
        },
      },
    },
    responses: {
      201: {
        description: 'Successful png creation',
        content: {
          'image/png': {
            schema: {
              type: 'string',
              format: 'binary',
            },
          },
        },
      },
    },
  },
};

module.exports = pngCreate;
