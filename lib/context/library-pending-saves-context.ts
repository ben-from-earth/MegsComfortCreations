import { createContext } from 'react';

export type RegisterPendingSave = (
  blockID: string,
  save: () => Promise<void>,
) => () => void;

const LibraryPendingSavesContext = createContext<RegisterPendingSave>(
  () => () => {},
);

export default LibraryPendingSavesContext;
