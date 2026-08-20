import {
  COLLECTOR_ITEM_SCHEMA_FAILED_MESSAGE,
  toCollectorSubmitErrorMessage,
} from '@/mediacollector/collector-submit-error-display';

describe('toCollectorSubmitErrorMessage', () => {
  test('uses the PNG format field message', () => {
    expect(
      toCollectorSubmitErrorMessage({
        pngFormat: {
          type: 'required',
          message: 'Please select a PNG template option',
        },
      }),
    ).toBe('Please select a PNG template option');
  });

  test('uses the book club field message', () => {
    expect(
      toCollectorSubmitErrorMessage({
        bookClubRepeat: {
          type: 'min',
          message: 'Book Club Repeat Number must be at least 1.',
        },
      }),
    ).toBe('Book Club Repeat Number must be at least 1.');
  });

  test('maps collected item schema failures', () => {
    expect(
      toCollectorSubmitErrorMessage({
        collectedData: [{ images: { type: 'too_small', message: 'Too small' } }],
      }),
    ).toBe(COLLECTOR_ITEM_SCHEMA_FAILED_MESSAGE);
  });
});
