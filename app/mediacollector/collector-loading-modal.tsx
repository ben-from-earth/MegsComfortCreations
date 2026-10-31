import CircularProgress from '@mui/material/CircularProgress';
import CheckIcon from '@mui/icons-material/Check';
import { MEDIA_TYPE_DEFINITIONS } from 'lib/constants/media-types';
import type { CollectorFormData } from './collector-form/collector-form-schema';
import { type CollectorSearchProgress } from './collector-search-progress';

const stepCircleClassName =
  'flex h-5 w-5 shrink-0 items-center justify-center rounded-full text-[11px] font-extrabold';

function LoadingCard({ children }: { children: React.ReactNode }) {
  return (
    <div className="fixed inset-0 z-100 flex items-center justify-center bg-[rgba(30,30,25,0.35)] p-6">
      <div
        role="status"
        className="flex min-h-41 w-85 max-w-full flex-col gap-4 rounded-[20px] bg-white px-4.5 py-4 text-[#2b2826] shadow-[0_24px_60px_rgba(30,30,25,0.3)]"
      >
        {children}
      </div>
    </div>
  );
}

function LoadingHeading({
  title,
  subtitle,
}: {
  title: string;
  subtitle: string;
}) {
  return (
    <div className="flex items-center gap-3.5">
      <CircularProgress size={28} thickness={4} sx={{ color: '#b3aaa0' }} />
      <div className="flex min-w-0 flex-col gap-0.5">
        <p className="m-0 text-lg leading-tight font-extrabold">{title}</p>
        <p className="m-0 truncate text-xs font-semibold text-[#6b645d]">
          {subtitle}
        </p>
      </div>
    </div>
  );
}

function pluralizeItems(count: number) {
  return `${count} ${count === 1 ? 'item' : 'items'}`;
}

export function CollectingCoversModal({
  progress,
  collectionList,
}: {
  progress: CollectorSearchProgress;
  collectionList: CollectorFormData['collectionList'];
}) {
  const collectFillPercent =
    progress.phase === 'pickingColors'
      ? 100
      : progress.totalCount <= 0
        ? 0
        : (100 * Math.min(progress.searchedCount, progress.totalCount)) /
          progress.totalCount;
  const colorFillPercent =
    progress.phase === 'pickingColors' && progress.isComplete ? 100 : 0;
  const isPickingColors = progress.phase === 'pickingColors';
  const mediaCountLabels = MEDIA_TYPE_DEFINITIONS.flatMap(
    ({ mediaType, label, plural }) => {
      const count = collectionList[mediaType].length;
      return count > 0
        ? [`${count} ${(count === 1 ? label : plural).toLowerCase()}`]
        : [];
    },
  );

  return (
    <LoadingCard>
      <div className="flex flex-wrap gap-1.5">
        {mediaCountLabels.map((mediaCountLabel) => (
          <span
            key={mediaCountLabel}
            className="rounded-full bg-[#f3e6df] px-2.5 py-0.75 text-[11px] font-bold text-[#5f4a40]"
          >
            {mediaCountLabel}
          </span>
        ))}
      </div>
      {progress.phase === 'collecting' ? (
        <LoadingHeading
          title="Collecting covers"
          subtitle={`${progress.searchedCount} of ${progress.totalCount} titles searched`}
        />
      ) : (
        <LoadingHeading
          title="Picking colors"
          subtitle="Reading spine colors from the covers"
        />
      )}
      <div className="mt-auto flex flex-col gap-1.5">
        <div className="flex gap-1">
          <div className="h-2 flex-4 overflow-hidden rounded-full bg-[#ece6de]">
            <div
              className="h-full rounded-full bg-[#6f7f66] transition-[width] duration-300"
              style={{ width: `${collectFillPercent}%` }}
            />
          </div>
          <div
            className={`h-2 flex-1 overflow-hidden rounded-full bg-[#ece6de] ${
              isPickingColors && colorFillPercent < 100 ? 'animate-pulse' : ''
            }`}
          >
            <div
              className="h-full rounded-full bg-[#6f7f66] transition-[width] duration-300"
              style={{ width: `${colorFillPercent}%` }}
            />
          </div>
        </div>
        <div className="flex justify-between text-[11px] font-bold">
          <span className={isPickingColors ? 'text-[#6b645d]' : ''}>
            Collecting covers
          </span>
          <span className={isPickingColors ? '' : 'text-[#6b645d]'}>
            Picking colors
          </span>
        </div>
      </div>
    </LoadingCard>
  );
}

export function SavingItemsModal({ newItemCount }: { newItemCount: number }) {
  return (
    <LoadingCard>
      <div className="flex items-center gap-2.5 text-xs font-bold">
        <span className={`${stepCircleClassName} bg-[#2a4637] text-white`}>
          1
        </span>
        <span>Save to database</span>
        <span className="h-px flex-1 bg-[#e0d9d0]" />
        <span
          className={`${stepCircleClassName} border border-[#b3aaa0] text-[#6b645d]`}
        >
          2
        </span>
        <span className="text-[#6b645d]">Build PNG</span>
      </div>
      <LoadingHeading
        title={`Saving ${newItemCount} new ${newItemCount === 1 ? 'item' : 'items'}`}
        subtitle="Adding them to the database"
      />
    </LoadingCard>
  );
}

export function BuildingPngModal({
  savedItemCount,
  customerName,
  orderNumber,
}: {
  savedItemCount: number;
  customerName: string;
  orderNumber: string;
}) {
  const orderDetails = [
    customerName.trim(),
    orderNumber.trim() ? `#${orderNumber.trim()}` : '',
  ]
    .filter(Boolean)
    .join(' · ');

  return (
    <LoadingCard>
      <div className="flex items-center gap-2.5 text-xs font-bold">
        <span className={`${stepCircleClassName} bg-[#e3e9df] text-[#3f5637]`}>
          <CheckIcon sx={{ fontSize: 13 }} />
        </span>
        <span className="text-[#6b645d]">
          Saved {pluralizeItems(savedItemCount)}
        </span>
        <span className="h-px flex-1 bg-[#6f7f66]" />
        <span className={`${stepCircleClassName} bg-[#2a4637] text-white`}>
          2
        </span>
        <span>Build PNG</span>
      </div>
      <LoadingHeading
        title="Building the PNG"
        subtitle={orderDetails || 'Putting the export together'}
      />
    </LoadingCard>
  );
}
