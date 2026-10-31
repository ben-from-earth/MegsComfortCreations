export type DefaultCoverRibbonVariant = 'saved' | 'new';

const ribbonFill: Record<DefaultCoverRibbonVariant, string> = {
  saved: '#2b2826',
  new: '#8a5a22',
};

const ribbonLabel: Record<DefaultCoverRibbonVariant, string> = {
  saved: 'Default cover',
  new: 'New default (unsaved)',
};

export function defaultCoverRibbonVariant(
  image: { url: string; isDefault?: boolean },
  originalDefaultUrl: string | undefined,
): DefaultCoverRibbonVariant | null {
  if (!image.isDefault) {
    return null;
  }
  return originalDefaultUrl === undefined || image.url === originalDefaultUrl
    ? 'saved'
    : 'new';
}

export function defaultCoverTileLabel(
  coverNumber: number,
  variant: DefaultCoverRibbonVariant | null,
) {
  const suffix =
    variant === 'saved'
      ? ', default'
      : variant === 'new'
        ? ', new default, unsaved'
        : '';
  return `Cover option ${coverNumber}${suffix}`;
}

export function DefaultCoverRibbon({
  variant,
}: {
  variant: DefaultCoverRibbonVariant;
}) {
  return (
    <svg
      viewBox="0 0 16 26"
      role="img"
      aria-label={ribbonLabel[variant]}
      className="pointer-events-none absolute top-0 left-2 h-6.5 w-4 drop-shadow-[0_1px_1.5px_rgba(0,0,0,0.35)]"
    >
      <path d="M0 0h16v26l-8-6-8 6z" fill={ribbonFill[variant]} />
    </svg>
  );
}

export function SetDefaultRibbonButton({
  coverNumber,
  onClick,
}: {
  coverNumber: number;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      aria-label={`Set cover ${coverNumber} as default`}
      title="Set as default"
      className="absolute top-0 left-2 z-20 h-6.5 w-4 cursor-pointer border-0 bg-transparent p-0"
    >
      <svg width="16" height="26" viewBox="0 0 16 26" aria-hidden="true">
        <path
          d="M1 1h14v23l-7-5.25L1 24z"
          fill="rgba(255,255,255,0.8)"
          stroke="#2b2826"
          strokeWidth="1.5"
          strokeDasharray="3 2"
        />
        <path
          d="M8 6v8M4 10h8"
          stroke="#2b2826"
          strokeWidth="1.8"
          strokeLinecap="round"
        />
      </svg>
    </button>
  );
}
