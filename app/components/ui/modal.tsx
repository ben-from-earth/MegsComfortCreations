import { useEffect, useRef, type ComponentProps } from 'react';
import { twMerge } from 'tailwind-merge';

export interface ModalProps
  extends Omit<ComponentProps<'dialog'>, 'ref' | 'open' | 'onCancel'> {
  onClose?: () => void;
}

export default function Modal({
  onClose,
  onKeyDown,
  className,
  children,
  ...dialogProps
}: ModalProps) {
  const dialogRef = useRef<HTMLDialogElement>(null);

  useEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) {
      return;
    }
    const opener = document.activeElement;
    dialog.showModal();
    return () => {
      dialog.close();
      if (opener instanceof HTMLElement) {
        opener.focus();
      }
    };
  }, []);

  return (
    <dialog
      ref={dialogRef}
      onKeyDown={(event) => {
        onKeyDown?.(event);
        // Escape must be stopped at keydown: browsers ignore preventDefault on
        // a repeated `cancel` event and would close the dialog behind React.
        if (event.key === 'Escape') {
          event.preventDefault();
          event.stopPropagation();
          onClose?.();
        }
      }}
      onCancel={(event) => event.preventDefault()}
      className={twMerge(
        'fixed inset-0 m-0 hidden h-full max-h-none w-full max-w-none items-center justify-center border-0 bg-[rgba(43,48,40,0.45)] p-4 text-[#2b2826] backdrop-blur-sm backdrop:bg-transparent open:flex',
        className,
      )}
      {...dialogProps}
    >
      {children}
    </dialog>
  );
}
