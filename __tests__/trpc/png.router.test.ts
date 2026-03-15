import { outputAuto } from 'lib/helpers/outputPNG';
import { createAdminContext, createTrpcCaller } from '../helpers/trpcTestContext';

jest.mock('lib/helpers/outputPNG', () => ({
  outputAuto: jest.fn(),
}));

describe('png router', () => {
  test('create delegates to outputAuto and returns base64 payload', async () => {
    const mockedOutputAuto = outputAuto as jest.Mock;
    mockedOutputAuto.mockResolvedValueOnce({
      mime: 'image/png',
      filename: 'Jane_D_123.png',
      buffer: Buffer.from('file-bytes'),
    });

    const caller = createTrpcCaller(createAdminContext({}));
    const response = await caller.png.create({
      template: 3,
      repeatCount: 2,
      customerName: 'Jane Doe',
      orderNumber: '123',
      images: [
        { url: 'https://img/1.png', type: 'book', spineColor: '#111111' },
      ],
    });

    expect(mockedOutputAuto).toHaveBeenCalledWith(
      expect.objectContaining({
        template: 3,
        fileOutputName: 'Jane_D_123',
      }),
    );
    expect(response).toEqual({
      mime: 'image/png',
      filename: 'Jane_D_123.png',
      dataBase64: Buffer.from('file-bytes').toString('base64'),
    });
  });

  test('create validates repeatCount minimum via zod', async () => {
    const caller = createTrpcCaller(createAdminContext({}));
    await expect(
      caller.png.create({
        template: 3,
        repeatCount: 0,
        customerName: 'Jane Doe',
        orderNumber: '123',
        images: [
          { url: 'https://img/1.png', type: 'book', spineColor: '#111111' },
        ],
      }),
    ).rejects.toMatchObject({
      code: 'BAD_REQUEST',
    });
  });
});
