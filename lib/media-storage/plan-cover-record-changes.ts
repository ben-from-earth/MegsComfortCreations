export type CoverRecordPlan = {
  updates: Array<{ id: string; incomingIndex: number }>;
  inserts: Array<{ incomingIndex: number }>;
  deletes: string[];
};

type IncomingCover = {
  id?: string | null;
};

function incomingCoverId(cover: IncomingCover) {
  if (typeof cover.id !== 'string' || cover.id.length === 0) {
    return null;
  }
  return cover.id;
}

export function planCoverRecordChanges(
  existingIds: readonly string[],
  incoming: readonly IncomingCover[],
): CoverRecordPlan {
  const existingIdSet = new Set(existingIds);
  const claimedIds = new Set<string>();
  const updates: CoverRecordPlan['updates'] = [];
  const inserts: CoverRecordPlan['inserts'] = [];

  for (let incomingIndex = 0; incomingIndex < incoming.length; incomingIndex++) {
    const id = incomingCoverId(incoming[incomingIndex]);
    if (id !== null && existingIdSet.has(id) && !claimedIds.has(id)) {
      claimedIds.add(id);
      updates.push({ id, incomingIndex });
      continue;
    }
    inserts.push({ incomingIndex });
  }

  return {
    updates,
    inserts,
    deletes: existingIds.filter((id) => !claimedIds.has(id)),
  };
}
