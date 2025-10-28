export const updateQueryCount = (): void => {
  const today: string = new Date().toISOString().split("T")[0];
  const storedDate: string | null = localStorage.getItem("lastQueryDate");

  if (storedDate !== today) {
    localStorage.setItem("queryCount", "0");
    localStorage.setItem("lastQueryDate", today);
  }

  let qCount = Number(localStorage.getItem("queryCount"));
  qCount++;
  localStorage.setItem("queryCount", `${qCount}`);
};
