/** @jest-environment jsdom */
import React from 'react';
import { FormProvider, useForm } from 'react-hook-form';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import CBBImages from '@/mediacollector/CBBImages';
import type { CollectorFormData } from '@/mediacollector/collector-form/collectorFormSchema';

const mockUploadCoverImage = jest.fn();

jest.mock('lib/trpc/client', () => ({
  trpc: {
    collect: {
      uploadCoverImage: {
        useMutation: () => ({
          mutateAsync: mockUploadCoverImage,
        }),
      },
    },
  },
}));

jest.mock('next/image', () => ({
  __esModule: true,
  default: (props: Record<string, unknown>) => {
    const { alt, onError, fill, unoptimized, loader, ...rest } = props;
    void fill;
    void unoptimized;
    void loader;
    const imageErrorHandler =
      typeof onError === 'function'
        ? (onError as React.ReactEventHandler<HTMLImageElement>)
        : undefined;
    return <img alt={String(alt ?? '')} onError={imageErrorHandler} {...rest} />;
  },
}));

function CBBImagesUploadHarness({
  mediaType = 'book',
}: {
  mediaType?: 'book' | 'movie' | 'videoGame' | 'album';
}) {
  const formMethods = useForm<CollectorFormData>({
    defaultValues: {
      orderNumber: '123',
      customerName: 'Test Customer',
      bookClubRepeat: 0,
      collectionList: {
        book: [],
        movie: [],
        videoGame: [],
        album: [],
      },
      collectedData: [
        {
          type: mediaType,
          images: [
            {
              url: 'https://img/first.png',
              selected: true,
              isDefault: true,
              spineColor: '#ffffff',
            },
          ],
          blockInfo: {
            title: 'Dune',
            author: 'Frank Herbert',
            pubYear: 1965,
            pageCount: 412,
            spineColor: '#ffffff',
            genres: [],
          },
          blockID: 'BLK-1',
          isDatabase: false,
        },
      ],
      pngFormat: '3',
    },
  });

  return (
    <FormProvider {...formMethods}>
      <CBBImages blockID={0} spineColor="#ffffff" />
    </FormProvider>
  );
}

describe('CBBImages upload slot', () => {
  beforeEach(() => {
    mockUploadCoverImage.mockReset();
  });

  test('shows upload placeholder for non-database book blocks only', () => {
    const { unmount } = render(<CBBImagesUploadHarness mediaType="book" />);
    expect(screen.getByLabelText('Add uploaded book image')).toBeTruthy();

    unmount();
    render(<CBBImagesUploadHarness mediaType="movie" />);
    expect(screen.queryByLabelText('Add uploaded book image')).toBeNull();
  });

  test('appends uploaded image, selects it, and hides placeholder', async () => {
    mockUploadCoverImage.mockResolvedValueOnce({
      url: 'https://cdn.example.com/custom-cover.png',
      selected: true,
      isDefault: false,
      spineColor: '#ffffff',
    });

    render(<CBBImagesUploadHarness mediaType="book" />);
    const input = screen.getByLabelText('Upload book image');
    const file = new File(['image-bytes'], 'cover.png', { type: 'image/png' });
    fireEvent.change(input, { target: { files: [file] } });

    await waitFor(() => {
      expect(mockUploadCoverImage).toHaveBeenCalledWith(
        expect.objectContaining({
          blockID: 'BLK-1',
          fileName: 'cover.png',
          mimeType: 'image/png',
        }),
      );
    });

    await waitFor(() => {
      expect(screen.queryByLabelText('Add uploaded book image')).toBeNull();
    });
    await waitFor(() => {
      expect(screen.getAllByAltText('book image')).toHaveLength(2);
    });
    expect(screen.getAllByText('Selected').length).toBeGreaterThan(0);
  });
});
