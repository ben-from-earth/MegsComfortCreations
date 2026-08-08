// library imports
import { twMerge } from 'tailwind-merge';
import React from 'react';

// components
import Button from '@/components/ui/Button';

export interface DialogProps {
  title: string;
  onClose: () => void;
  children: React.ReactNode;
  className?: string;
}

function DialogHeader({
  title,
  onClose,
}: {
  title: string;
  onClose: () => void;
}) {
  return (
    <div className="bg-lightpink border-darkpink mb-0 flex w-full items-center gap-4 border-3 p-2">
      <h2 className="min-w-0 flex-1 text-left text-4xl tracking-wider">
        {title}
      </h2>
      <Button
        variant="primary"
        className="ml-auto shrink-0"
        onClick={onClose}
        label="Close"
        width={100}
        fontSize={25}
      />
    </div>
  );
}

export default function Dialog({
  title,
  onClose,
  children,
  className,
}: DialogProps) {
  return (
    <>
      <div
        aria-hidden="true"
        data-testid="dialog-backdrop"
        className="fixed inset-0 z-90 bg-black/20 backdrop-blur-sm"
      />
      <div
        role="dialog"
        aria-modal="true"
        aria-label={title}
        className={twMerge(
          'fixed top-1/2 left-1/2 z-100 flex max-h-100 w-fit max-w-[90vw] -translate-x-1/2 -translate-y-1/2 flex-col content-center items-center overflow-y-auto text-4xl tracking-wider text-black',
          className,
        )}
      >
        <DialogHeader title={title} onClose={onClose} />
        <div className="w-full rounded-b-md border-3 border-t-0 border-[#a8a68d] bg-[#f1efd3] px-10 pt-2 pb-4">
          {children}
        </div>
      </div>
    </>
  );
}
