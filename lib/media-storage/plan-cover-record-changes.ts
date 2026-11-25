export type CoverRecordPlan = {
  updates: Array<{ id: string; incomingIndex: number }>;
  inserts: Array<{ incomingIndex: number }>;
  deletes: string[];
};

export function planCoverRecordChanges(
  existingIds: readonly string[],
  incoming: readonly { id?: string }[],
): CoverRecordPlan {
  const existingIdSet = new Set(existingIds);
  const claimedIds = new Set<string>();
  const updates: CoverRecordPlan['updates'] = [];
  const inserts: CoverRecordPlan['inserts'] = [];

  incoming.forEach(({ id }, incomingIndex) => {
    if (id && existingIdSet.has(id) && !claimedIds.has(id)) {
      claimedIds.add(id);
      updates.push({ id, incomingIndex });
      return;
    }
    inserts.push({ incomingIndex });
  });

  return {
    updates,
    inserts,
    deletes: existingIds.filter((id) => !claimedIds.has(id)),
  };
}
