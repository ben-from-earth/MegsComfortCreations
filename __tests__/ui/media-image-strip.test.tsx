/** @jest-environment jsdom */
import React from 'react';
import { fireEvent, render, screen } from '@testing-library/react';
import MediaImageStrip from '@/shared/MediaImageStrip';

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

describe('MediaImageStrip', () => {
  test('renders default badge and overflow button for extra covers', () => {
    render(
      <MediaImageStrip
        mediaType="book"
        images={[
          { url: 'https://img/1.png', isDefault: true, selected: true, spineColor: '#111111' },
          { url: 'https://img/2.png', isDefault: false, selected: false, spineColor: '#222222' },
          { url: 'https://img/3.png', isDefault: false, selected: false, spineColor: '#333333' },
          { url: 'https://img/4.png', isDefault: false, selected: false, spineColor: '#444444' },
          { url: 'https://img/5.png', isDefault: false, selected: false, spineColor: '#555555' },
        ]}
      />,
    );

    expect(screen.getAllByTestId('StarsIcon').length).toBeGreaterThan(0);
    expect(screen.getByRole('button', { name: '+1' })).toBeTruthy();
  });

  test('uses selected image as primary and maps overflow click index', () => {
    const onImageClick = jest.fn();
    render(
      <MediaImageStrip
        mediaType="book"
        onImageClick={onImageClick}
        images={[
          { url: 'https://img/1.png', isDefault: true, selected: false, spineColor: '#111111' },
          { url: 'https://img/2.png', isDefault: false, selected: true, spineColor: '#222222' },
          { url: 'https://img/3.png', isDefault: false, selected: false, spineColor: '#333333' },
          { url: 'https://img/4.png', isDefault: false, selected: false, spineColor: '#444444' },
          { url: 'https://img/5.png', isDefault: false, selected: false, spineColor: '#555555' },
          { url: 'https://img/6.png', isDefault: false, selected: false, spineColor: '#666666' },
        ]}
      />,
    );

    fireEvent.click(screen.getByRole('button', { name: '+2' }));
    const popoverImages = screen.getAllByAltText('book image');
    fireEvent.click(popoverImages[popoverImages.length - 1]);

    expect(onImageClick).toHaveBeenCalledWith(5);
  });
});
