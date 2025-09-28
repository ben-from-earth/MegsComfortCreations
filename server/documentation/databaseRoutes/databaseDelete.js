const databaseDelete = {
  delete: {
    tags: ['Database'],
    summary: 'Delete media item from the database',
    description:
      'Using type and title, search the corresponding type table for the row with corresponding title and delete it',
    parameters: [
      {
        in: 'query',
        name: 'type',
        required: true,
        description: 'Type of media being searched and deleted',
        schema: {
          type: 'string',
        },
      },
      {
        in: 'query',
        name: 'title',
        required: true,
        description: 'Title of media being searched and deleted',
        schema: {
          type: 'string',
        },
      },
    ],
    responses: {
      200: {
        description:
          'Found item in the database with type/title combo and successfully deleted',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                message: { type: 'string', description: 'Success message' },
              },
            },
            example: {
              message: 'Successfully deleted [title]',
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
                  description: 'Non-existent deletion request',
                },
                message: {
                  type: 'string',
                  description:
                    'No item with title: [title] in the [type] database exists',
                },
              },
            },
            example: {
              error: 'Non-existent deletion request',
              message:
                'No item with title: [title] in the [type] database exists',
            },
          },
        },
      },
      400: {
        description: 'Deletion error',
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
              error: 'Database deletion error',
              message: 'Error deleting items from the database.',
            },
          },
        },
      },
    },
  },
};

module.exports = databaseDelete;
