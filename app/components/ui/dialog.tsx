import { twMerge } from 'tailwind-merge';
import React from 'react';
import Button from '@/components/ui/button';
import Modal from '@/components/ui/modal';

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
    <Modal aria-label={title} onClose={onClose}>
      <div
        className={twMerge(
          'flex max-h-100 w-fit max-w-[90vw] flex-col content-center items-center overflow-y-auto text-4xl tracking-wider text-black',
          className,
        )}
      >
        <DialogHeader title={title} onClose={onClose} />
        <div className="w-full rounded-b-md border-3 border-t-0 border-[#a8a68d] bg-[#f1efd3] px-10 pt-2 pb-4">
          {children}
        </div>
      </div>
    </Modal>
  );
}
