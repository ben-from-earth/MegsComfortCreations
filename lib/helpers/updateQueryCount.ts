export const updateQueryCount = (): void => {
  if (typeof window === 'undefined' || !('localStorage' in globalThis)) {
    return;
  }

  const today: string = new Date().toISOString().split('T')[0];
  const storedDate: string | null =
    window.localStorage.getItem('lastQueryDate');

  if (storedDate !== today) {
    window.localStorage.setItem('queryCount', '0');
    window.localStorage.setItem('lastQueryDate', today);
  }

  const current = window.localStorage.getItem('queryCount');
  const qCount = Number(current ?? '0') + 1;
  window.localStorage.setItem('queryCount', String(qCount));
};
