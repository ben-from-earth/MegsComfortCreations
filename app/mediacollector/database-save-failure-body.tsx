// interfaces and types
import type { DatabaseSaveFailureDisplayLine } from './database-save-error-display';

export default function DatabaseSaveFailureBody({
  failureLines,
}: {
  failureLines: DatabaseSaveFailureDisplayLine[];
}) {
  const failedCount = failureLines.length;
  const titleWord = failedCount === 1 ? 'title' : 'titles';

  return (
    <>
      <p>
        {failedCount} {titleWord} experienced errors when attempting to save to
        the database. All blocks besides the following were successfully saved:
      </p>
      <ol className="mt-4 list-decimal pl-10">
        {failureLines.map((line) => (
          <li
            key={line.blockID || `${line.title}-${line.blockNumber}`}
            className="mb-2"
          >
            {line.blockNumber !== null
              ? `${line.title} in Block #${line.blockNumber}: ${line.reason}`
              : `${line.title}: ${line.reason}`}
          </li>
        ))}
      </ol>
    </>
  );
}
