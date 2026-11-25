import { planCoverRecordChanges } from 'lib/media-storage/plan-cover-record-changes';

describe('planCoverRecordChanges', () => {
  test('updates a cover on this parent, inserts a new cover, and deletes a missing one', () => {
    expect(
      planCoverRecordChanges(['keep', 'drop'], [{ id: 'keep' }, {}]),
    ).toEqual({
      updates: [{ id: 'keep', incomingIndex: 0 }],
      inserts: [{ incomingIndex: 1 }],
      deletes: ['drop'],
    });
  });

  test('inserts an id that belongs to another parent', () => {
    expect(
      planCoverRecordChanges(['cover-a'], [{ id: 'cover-from-elsewhere' }]),
    ).toEqual({
      updates: [],
      inserts: [{ incomingIndex: 0 }],
      deletes: ['cover-a'],
    });
  });

  test('uses an id only once', () => {
    expect(
      planCoverRecordChanges(['keep'], [{ id: 'keep' }, { id: 'keep' }]),
    ).toEqual({
      updates: [{ id: 'keep', incomingIndex: 0 }],
      inserts: [{ incomingIndex: 1 }],
      deletes: [],
    });
  });
});
