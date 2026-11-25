import Link from 'next/link';
import {
  MEDIA_TYPE_DEFINITIONS,
  type MediaType,
} from 'lib/constants/media-types';

export type OrderTableRow = {
  id: string;
  customerName: string;
  orderNumber: string;
  template: number;
  repeatCount: number;
  fileUrl: string;
  fileName: string;
  updatedAt: Date;
  counts: Record<MediaType, number>;
  firstCover: { url: string } | null;
};

const updatedThisYear = new Intl.DateTimeFormat('en-US', {
  month: 'short',
  day: 'numeric',
});
const updatedWithYear = new Intl.DateTimeFormat('en-US', {
  month: 'short',
  day: 'numeric',
  year: 'numeric',
});

function formatUpdated(date: Date) {
  const format =
    date.getFullYear() === new Date().getFullYear()
      ? updatedThisYear
      : updatedWithYear;
  return format.format(date);
}

function countSummary(counts: Record<MediaType, number>) {
  const parts = MEDIA_TYPE_DEFINITIONS.flatMap((definition) => {
    const count = counts[definition.mediaType];
    if (count === 0) {
      return [];
    }
    const label = count === 1 ? definition.label : definition.plural;
    return [`${count} ${label.toLowerCase()}`];
  });
  return parts.length > 0 ? parts.join(' · ') : 'No covers';
}

const headerCellClassName = 'border-b border-[#efe9e2] px-3 py-3.5';
const actionClassName =
  'inline-flex h-9 items-center rounded-lg border border-[#ddd5cb] bg-white px-3 text-[13px] font-extrabold text-[#2b2826] no-underline';

export default function OrdersTable({
  rows,
  emptyDatabase,
}: {
  rows: OrderTableRow[];
  emptyDatabase: boolean;
}) {
  return (
    <div className="overflow-x-auto rounded-2xl border border-[#e6e0d8] bg-white">
      {rows.length === 0 ? (
        <p className="m-0 px-6 py-16 text-center text-sm font-semibold text-[#6b645d]">
          {emptyDatabase
            ? 'No orders yet. Exporting from the collector saves one.'
            : 'No orders match these filters.'}
        </p>
      ) : (
        <table className="w-full min-w-[960px] border-collapse text-sm">
          <thead>
            <tr className="text-left text-[11px] font-extrabold tracking-[0.6px] text-[#6b645d] uppercase">
              <th className="border-b border-[#efe9e2] px-[18px] py-3.5">
                Order
              </th>
              <th className={headerCellClassName}>Covers</th>
              <th className={headerCellClassName}>Template</th>
              <th className={headerCellClassName}>Repeat</th>
              <th className={headerCellClassName}>Updated</th>
              <th className="border-b border-[#efe9e2] px-[18px] py-3.5">
                <span className="sr-only">Actions</span>
              </th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr key={row.id} className="border-b border-[#f3efe9]">
                <td className="px-[18px] py-2.5">
                  <div className="flex items-center gap-3.5">
                    {row.firstCover ? (
                      <img
                        src={row.firstCover.url}
                        alt=""
                        className="h-[52px] w-[52px] rounded-lg object-cover"
                      />
                    ) : (
                      <span
                        aria-hidden="true"
                        className="h-[52px] w-[52px] rounded-lg bg-[#efebe5]"
                      />
                    )}
                    <span className="flex min-w-0 flex-col gap-0.5">
                      <span className="font-extrabold text-[#2b2826]">
                        {row.customerName}
                      </span>
                      <span className="text-xs font-semibold text-[#6b645d]">
                        #{row.orderNumber}
                      </span>
                    </span>
                  </div>
                </td>
                <td className="px-3 py-2.5 font-semibold text-[#4a443f]">
                  {countSummary(row.counts)}
                </td>
                <td className="px-3 py-2.5 font-semibold whitespace-nowrap text-[#4a443f]">
                  {row.template} mm
                </td>
                <td className="px-3 py-2.5 font-semibold whitespace-nowrap text-[#4a443f]">
                  ×{row.repeatCount}
                </td>
                <td className="px-3 py-2.5 text-[13px] font-semibold whitespace-nowrap text-[#6b645d]">
                  {formatUpdated(row.updatedAt)}
                </td>
                <td className="px-[18px] py-2.5 text-right whitespace-nowrap">
                  <a
                    href={row.fileUrl}
                    download={row.fileName}
                    className={actionClassName}
                  >
                    Download
                  </a>
                  <Link
                    href={`/mediacollector?order=${row.id}`}
                    className={`${actionClassName} ml-1`}
                  >
                    Open in collector
                  </Link>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}
