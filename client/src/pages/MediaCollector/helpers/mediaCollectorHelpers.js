const updateQueryCount = () => {
  const today = new Date().toISOString().split("T")[0];
  const storedDate = localStorage.getItem("lastQueryDate");

  if (storedDate !== today) {
    localStorage.setItem("queryCount", "0");
    localStorage.setItem("lastQueryDate", today);
  }

  let qCount = Number(localStorage.getItem("queryCount"));
  qCount++;
  localStorage.setItem("queryCount", `${qCount}`);
};

const titleRearrange = (title) => {
  if (title.startsWith("The ")) {
    const newTitle = title.slice(4) + ", The";
    return newTitle;
  }

  if (title.startsWith("A ")) {
    const newTitle = title.slice(2) + ", A";
    return newTitle;
  }

  if (title.startsWith("An ")) {
    const newTitle = title.slice(3) + ", An";
    return newTitle;
  }

  if (title.endsWith(", The")) {
    const newTitle = "The " + title.slice(0, title.length - 5);
    return newTitle;
  }

  if (title.endsWith(", A")) {
    const newTitle = "A " + title.slice(0, title.length - 3);
    return newTitle;
  }

  if (title.endsWith(", An")) {
    const newTitle = "The " + title.slice(0, title.length - 4);
    return newTitle;
  }
  return title;
};

export { updateQueryCount, titleRearrange };
