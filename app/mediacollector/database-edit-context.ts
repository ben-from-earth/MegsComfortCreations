import { createContext } from 'react';
import type { MediaItemForm } from './collector-form/media-item-form-schema';

export interface DatabaseEditContextValue {
  originalsByBlockId: ReadonlyMap<string, MediaItemForm>;
  unlockedBlockIds: ReadonlySet<string>;
  setBlockUnlocked: (blockID: string, isUnlocked: boolean) => void;
}

const DatabaseEditContext = createContext<DatabaseEditContextValue>({
  originalsByBlockId: new Map(),
  unlockedBlockIds: new Set(),
  setBlockUnlocked: () => {},
});

export default DatabaseEditContext;
