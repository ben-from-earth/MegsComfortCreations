import { useEffect, useId, useRef } from 'react';
import DeleteIcon from '@mui/icons-material/DeleteOutline';
import Modal from '@/components/ui/modal';

const buttonClassName =
  'inline-flex h-11 cursor-pointer items-center gap-2 rounded-[10px] px-4.5 font-extrabold focus-visible:outline-3 focus-visible:outline-offset-2 focus-visible:outline-[#3d3935] disabled:cursor-not-allowed disabled:opacity-60';

export default function OrderDeleteDialog({
  customerName,
  orderNumber,
  isDeleting,
  error,
  onCancel,
  onConfirm,
}: {
  customerName: string;
  orderNumber: string;
  isDeleting: boolean;
  error: string | null;
  onCancel: () => void;
  onConfirm: () => void;
}) {
  const cancelButtonRef = useRef<HTMLButtonElement>(null);
  const headingId = useId();
  const descriptionId = useId();

  useEffect(() => {
    cancelButtonRef.current?.focus();
  }, []);

  const handleCancel = () => {
    if (!isDeleting) {
      onCancel();
    }
  };

  return (
    <Modal
      role="alertdialog"
      aria-labelledby={headingId}
      aria-describedby={descriptionId}
      onClose={handleCancel}
      onClick={(event) => {
        if (event.target === event.currentTarget) {
          handleCancel();
        }
      }}
    >
      <div className="flex w-115 max-w-full flex-col overflow-hidden rounded-[20px] bg-white shadow-[0_24px_60px_rgba(30,30,25,0.3)]">
        <div className="flex flex-col gap-4.5 px-6.5 pt-6.5 pb-5">
          <div className="flex items-center gap-3.5">
            <span
              aria-hidden="true"
              className="flex h-11 w-11 shrink-0 items-center justify-center rounded-full bg-[#f7e3dc] text-[#9a3b26]"
            >
              <DeleteIcon sx={{ fontSize: 22 }} />
            </span>
            <h2
              id={headingId}
              className="m-0 text-[21px] leading-tight font-extrabold"
            >
              Delete this order?
            </h2>
          </div>

          <div className="flex min-w-0 flex-col gap-0.5 rounded-xl bg-[#f7f4f0] p-3">
            <span className="text-base font-extrabold">{customerName}</span>
            <span className="text-[13px] font-semibold text-[#6b645d]">
              #{orderNumber}
            </span>
          </div>

          <p
            id={descriptionId}
            className="m-0 text-sm leading-normal text-[#4a443f]"
          >
            This removes the order and its print file. Titles and covers in the
            library stay.{' '}
            <strong className="text-[#2b2826]">This can’t be undone.</strong>
          </p>
          {error ? (
            <p role="alert" className="m-0 text-[13px] font-bold text-[#9a3b26]">
              {error}
            </p>
          ) : null}
        </div>

        <div className="flex justify-end gap-2.5 border-t border-[#ece6de] bg-[#f3f6f0] px-5 py-3.5">
          <button
            ref={cancelButtonRef}
            type="button"
            disabled={isDeleting}
            onClick={handleCancel}
            className={`${buttonClassName} border border-[#ddd5cb] bg-white text-[#2b2826]`}
          >
            Cancel
          </button>
          <button
            type="button"
            disabled={isDeleting}
            onClick={onConfirm}
            className={`${buttonClassName} border-0 bg-[#9a3b26] text-white`}
          >
            <DeleteIcon sx={{ fontSize: 16 }} />
            {isDeleting ? 'Deleting…' : 'Delete order'}
          </button>
        </div>
      </div>
    </Modal>
  );
}
