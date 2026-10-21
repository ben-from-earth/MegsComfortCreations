import {
  BOOK_CLUB_REPEAT_MAX,
  collectorFormSchema,
  PNG_FORMAT_OPTIONS,
  PNG_FORMAT_VALUES,
  type CollectorFormData,
} from '@/mediacollector/collector-form/collector-form-schema';

const validCollectorForm = {
  orderNumber: '1001',
  customerName: 'Ada Lovelace',
  bookClubRepeat: 1,
  collectionList: {
    book: [],
    movie: [],
    videoGame: [],
    album: [],
  },
  collectedData: [],
  pngFormat: '3',
} satisfies CollectorFormData;

describe('collectorFormSchema header and PNG fields', () => {
  test('PNG radio options match the schema enum', () => {
    expect(PNG_FORMAT_OPTIONS.map((option) => option.value)).toEqual([
      ...PNG_FORMAT_VALUES,
    ]);
  });

  test('accepts a selected PNG format', () => {
    expect(collectorFormSchema.parse(validCollectorForm).pngFormat).toBe('3');
  });

  test('requires a PNG format', () => {
    const result = collectorFormSchema.safeParse({
      ...validCollectorForm,
      pngFormat: null,
    });

    expect(result.success).toBe(false);
  });

  test('requires book club repeat to be at least 1', () => {
    const result = collectorFormSchema.safeParse({
      ...validCollectorForm,
      bookClubRepeat: 0,
    });

    expect(result.success).toBe(false);
  });

  test('rejects book club repeat above 25', () => {
    expect(
      collectorFormSchema.safeParse({
        ...validCollectorForm,
        bookClubRepeat: BOOK_CLUB_REPEAT_MAX + 1,
      }).success,
    ).toBe(false);
  });
});
