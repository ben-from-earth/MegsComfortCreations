/** @jest-environment jsdom */
import React from 'react';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import EditDatabaseBlock from '@/showdatabase/EditDatabaseBlock';
import GenreContext from 'lib/context/GenreContext';

const mockDatabaseEdit = jest.fn();
const mockUploadCoverImage = jest.fn();
const mockLinkGenres = jest.fn();
const mockUnlinkGenres = jest.fn();
const mockInvalidateGenres = jest.fn();
const mockHandleGetMedia = jest.fn().mockResolvedValue(undefined);

jest.mock('lib/context/DatabasePageContext', () => ({
  useDatabasePageContext: () => ({
    handleGetMedia: mockHandleGetMedia,
  }),
}));

jest.mock('lib/trpc/client', () => ({
  trpc: {
    database: {
      edit: {
        useMutation: () => ({
          mutateAsync: mockDatabaseEdit,
        }),
      },
    },
    collect: {
      uploadCoverImage: {
        useMutation: () => ({
          mutateAsync: mockUploadCoverImage,
        }),
      },
    },
    genres: {
      link: {
        useMutation: () => ({
          mutateAsync: mockLinkGenres,
        }),
      },
      unlink: {
        useMutation: () => ({
          mutateAsync: mockUnlinkGenres,
        }),
      },
    },
    useUtils: () => ({
      genres: {
        getForBook: {
          invalidate: mockInvalidateGenres,
        },
      },
    }),
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

describe('EditDatabaseBlock image uploads', () => {
  beforeEach(() => {
    mockDatabaseEdit.mockReset();
    mockUploadCoverImage.mockReset();
    mockLinkGenres.mockReset();
    mockUnlinkGenres.mockReset();
    mockInvalidateGenres.mockReset();
    mockHandleGetMedia.mockClear();
  });

  test('shows upload control for book edit block', () => {
    render(
      <GenreContext.Provider value={[]}>
        <EditDatabaseBlock
          info={{
            type: 'book',
            images: [
              {
                url: 'https://img/book-1.png',
                isDefault: true,
                selected: true,
                spineColor: '#ffffff',
              },
            ],
            blockInfo: {
              title: 'Dune',
              author: 'Frank Herbert',
              pubYear: 1965,
              pageCount: 412,
              spineColor: '#ffffff',
              initialGenres: [],
            },
            id: 'book-id-1',
            setEdit: jest.fn(),
          }}
        />
      </GenreContext.Provider>,
    );

    expect(screen.getByLabelText('Add uploaded database image')).toBeTruthy();
  });

  test('appends uploaded image and submits updated images', async () => {
    mockUploadCoverImage.mockResolvedValueOnce({
      url: 'https://cdn.example.com/book-uploaded.png',
      selected: true,
      isDefault: false,
      spineColor: '#ffffff',
    });
    mockDatabaseEdit.mockResolvedValueOnce({ message: 'Saved' });

    render(
      <GenreContext.Provider value={[]}>
        <EditDatabaseBlock
          info={{
            type: 'book',
            images: [
              {
                url: 'https://img/book-1.png',
                isDefault: true,
                selected: true,
                spineColor: '#ffffff',
              },
            ],
            blockInfo: {
              title: 'Dune',
              author: 'Frank Herbert',
              pubYear: 1965,
              pageCount: 412,
              spineColor: '#ffffff',
              initialGenres: [],
            },
            id: 'book-id-1',
            setEdit: jest.fn(),
          }}
        />
      </GenreContext.Provider>,
    );

    expect(screen.getAllByAltText('book image')).toHaveLength(1);

    const input = screen.getByLabelText('Upload database image');
    const file = new File(['image-bytes'], 'database-cover.png', {
      type: 'image/png',
    });
    fireEvent.change(input, { target: { files: [file] } });

    await waitFor(() => {
      expect(mockUploadCoverImage).toHaveBeenCalledWith(
        expect.objectContaining({
          blockID: 'book-id-1',
          fileName: 'database-cover.png',
          mimeType: 'image/png',
          sortOrder: 1,
        }),
      );
    });

    await waitFor(() => {
      expect(screen.getAllByAltText('book image')).toHaveLength(2);
    });

    fireEvent.click(screen.getByText('Submit Changes'));

    await waitFor(() => {
      expect(mockDatabaseEdit).toHaveBeenCalledWith(
        expect.objectContaining({
          type: 'book',
          item: expect.objectContaining({
            images: [
              expect.objectContaining({
                url: 'https://img/book-1.png',
              }),
              expect.objectContaining({
                url: 'https://cdn.example.com/book-uploaded.png',
              }),
            ],
          }),
        }),
      );
    });
  });
});
