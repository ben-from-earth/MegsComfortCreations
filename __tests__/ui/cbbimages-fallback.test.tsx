/** @jest-environment jsdom */
import React from 'react';
import { FormProvider, useForm } from 'react-hook-form';
import { fireEvent, render, screen } from '@testing-library/react';
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

function CBBImagesTestHarness() {
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
          type: 'book',
          images: [
            {
              url: '/uploads/covers/2026/03/example.png',
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
          isDatabase: true,
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

describe('CBBImages broken image fallback', () => {
  test('renders fixed-size fallback when image load fails', () => {
    const { container } = render(<CBBImagesTestHarness />);
    const image = screen.getByAltText('book image');
    fireEvent.error(image);

    expect(screen.getByText('Image path broken')).toBeTruthy();
    expect(container.querySelector('[class*="h-31"]')).toBeTruthy();
    expect(container.querySelector('[class*="w-21"]')).toBeTruthy();
  });
});
