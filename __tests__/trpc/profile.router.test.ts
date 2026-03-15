import { createAdminContext, createTrpcCaller } from '../helpers/trpcTestContext';
import {
  getImageMigrationStatus,
} from 'lib/media-storage/media-image-records';
import { runOneTimeLegacyImageMigration } from 'lib/media-storage/migrate-legacy-image-urls';

jest.mock('lib/media-storage/media-image-records', () => ({
  getImageMigrationStatus: jest.fn(),
}));
jest.mock('lib/media-storage/migrate-legacy-image-urls', () => ({
  migrateLegacyImageUrlsToLocalFiles: jest.fn(),
  runOneTimeLegacyImageMigration: jest.fn(),
}));

describe('profile router', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('get returns expected static profile contract', async () => {
    const caller = createTrpcCaller(createAdminContext({}));
    const profile = await caller.profile.get();

    expect(profile).toEqual({
      id: 123,
      firstName: 'Ben',
      lastName: 'Knox',
      email: 'example@email.com',
    });
  });

  test('getImageMigrationStatus returns migration status payload', async () => {
    (getImageMigrationStatus as jest.Mock).mockResolvedValueOnce({
      totalItems: 5,
      externalUrlCount: 2,
      missingReferenceCount: 1,
      isCompleted: false,
    });

    const caller = createTrpcCaller(createAdminContext({}));
    const status = await caller.profile.getImageMigrationStatus();

    expect(status).toEqual({
      totalItems: 5,
      externalUrlCount: 2,
      missingReferenceCount: 1,
      isCompleted: false,
    });
  });

  test('migrateImageFiles executes one-time migration action', async () => {
    (runOneTimeLegacyImageMigration as jest.Mock).mockResolvedValueOnce({
      alreadyCompleted: false,
      statusBefore: {
        totalItems: 5,
        externalUrlCount: 2,
        missingReferenceCount: 0,
        isCompleted: false,
      },
      statusAfter: {
        totalItems: 5,
        externalUrlCount: 0,
        missingReferenceCount: 0,
        isCompleted: true,
      },
      summary: {
        dryRun: false,
        processedItems: 5,
        migratedExternalUrls: 2,
        skippedLocalPaths: 3,
        failedDownloads: 0,
        deletedRows: 0,
        failures: [],
      },
    });

    const caller = createTrpcCaller(createAdminContext({}));
    const response = await caller.profile.migrateImageFiles();

    expect(response).toMatchObject({
      dryRun: false,
      alreadyCompleted: false,
      summary: { migratedExternalUrls: 2 },
    });
    expect(runOneTimeLegacyImageMigration).toHaveBeenCalledTimes(1);
  });
});
