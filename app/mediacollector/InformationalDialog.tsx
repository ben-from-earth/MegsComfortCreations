// components
import Button from '@/components/ui/Button';

// interfaces and types
import type { DatabaseSaveFailureDisplayLine } from './database-save-error-display';

export interface InformationalDialogProps {
  variant: 'databaseSave' | 'informationalOnly';
  failureLines?: DatabaseSaveFailureDisplayLine[];
  infoText?: string;
  close: () => void;
}

export default function InformationalDialog({
  variant,
  failureLines,
  infoText,
  close,
}: InformationalDialogProps) {
  if (variant === 'databaseSave') {
    const lines = failureLines ?? [];
    const failedCount = lines.length;
    const titleWord = failedCount === 1 ? 'title' : 'titles';

    return (
      <div className="border-darkpink bg-lightpink fixed top-1/2 left-1/2 z-100 flex max-h-100 w-fit max-w-[90vw] -translate-x-1/2 -translate-y-1/2 flex-col content-center items-center overflow-y-auto rounded-md border-3 p-2 text-4xl tracking-wider text-black">
        <div className="mb-2 flex w-full">
          <Button
            variant="primary"
            className={'ml-auto'}
            onClick={close}
            label={'Close'}
            width={100}
            fontSize={25}
          />
        </div>

        <div className="px-10">
          <p>
            {failedCount} {titleWord} experienced errors when attempting to save
            to the database. All blocks besides the following were successfully
            saved:
          </p>
          <ol className="mt-4 list-decimal pl-10">
            {lines.map((line) => (
              <li key={line.blockID || `${line.title}-${line.blockNumber}`} className="mb-2">
                {line.blockNumber !== null
                  ? `${line.title} in Block #${line.blockNumber}: ${line.reason}`
                  : `${line.title}: ${line.reason}`}
              </li>
            ))}
          </ol>
        </div>
      </div>
    );
  } else if (variant === 'informationalOnly') {
    return (
      <div className="border-darkpink bg-lightpink fixed top-1/2 left-1/2 z-100 flex max-h-100 w-5/12 -translate-x-1/2 -translate-y-1/2 flex-col content-center items-center overflow-y-auto rounded-md border-3 px-6 py-2 text-4xl tracking-wider text-black">
        <div className="mb-2 flex w-full">
          <Button
            variant="primary"
            className={'ml-auto'}
            onClick={close}
            label={'Close'}
            width={100}
            fontSize={25}
          />
        </div>
        {infoText}
      </div>
    );
  }
}
