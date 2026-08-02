/** @jest-environment jsdom */
import React from 'react';
import { fireEvent, render, screen } from '@testing-library/react';
import InformationalDialog from '@/mediacollector/InformationalDialog';

describe('InformationalDialog databaseSave variant', () => {
  test('renders failure summary with block numbers and friendly reasons', () => {
    const close = jest.fn();

    render(
      <InformationalDialog
        variant="databaseSave"
        failureLines={[
          {
            blockID: 'BLK-2',
            title: 'Matrix',
            blockNumber: 2,
            reason:
              'The cover image could not be saved, so this item was not added to the database.',
          },
        ]}
        close={close}
      />,
    );

    expect(
      screen.getByText(
        /1 title experienced errors when attempting to save to the database/i,
      ),
    ).toBeTruthy();
    expect(
      screen.getByText(
        /All blocks besides the following were successfully saved/i,
      ),
    ).toBeTruthy();
    expect(
      screen.getByText(
        'Matrix in Block #2: The cover image could not be saved, so this item was not added to the database.',
      ),
    ).toBeTruthy();

    fireEvent.click(screen.getByRole('button', { name: 'Close' }));
    expect(close).toHaveBeenCalledTimes(1);
  });
});
