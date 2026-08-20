import type { FieldErrors } from 'react-hook-form';
import type { CollectorFormData } from './collector-form/collectorFormSchema';

export const COLLECTOR_ITEM_SCHEMA_FAILED_MESSAGE =
  'Some collected items have invalid details. Check the highlighted blocks.';

export const COLLECTOR_SUBMIT_FAILED_MESSAGE =
  'Please fix the highlighted fields and try again.';

export function toCollectorSubmitErrorMessage(
  errors: FieldErrors<CollectorFormData>,
): string | null {
  const pngFormatMessage = errors.pngFormat?.message;
  if (typeof pngFormatMessage === 'string') {
    return pngFormatMessage;
  }

  const bookClubRepeatMessage = errors.bookClubRepeat?.message;
  if (typeof bookClubRepeatMessage === 'string') {
    return bookClubRepeatMessage;
  }

  if (errors.collectedData) {
    return COLLECTOR_ITEM_SCHEMA_FAILED_MESSAGE;
  }

  return Object.keys(errors).length > 0
    ? COLLECTOR_SUBMIT_FAILED_MESSAGE
    : null;
}
