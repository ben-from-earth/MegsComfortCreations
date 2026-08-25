export const titleRearrange = (title: string): string => {
  if (title.toLowerCase().startsWith("the ")) {
    const newTitle: string = title.slice(4) + ", The";
    return newTitle;
  }

  if (title.toLowerCase().startsWith("a ")) {
    const newTitle: string = title.slice(2) + ", A";
    return newTitle;
  }

  if (title.toLowerCase().startsWith("an ")) {
    const newTitle: string = title.slice(3) + ", An";
    return newTitle;
  }

  if (title.toLowerCase().endsWith(", the")) {
    const newTitle: string = "The " + title.slice(0, title.length - 5);
    return newTitle;
  }

  if (title.toLowerCase().endsWith(", a")) {
    const newTitle: string = "A " + title.slice(0, title.length - 3);
    return newTitle;
  }

  if (title.toLowerCase().endsWith(", an")) {
    const newTitle: string = "An " + title.slice(0, title.length - 4);
    return newTitle;
  }
  return title;
};
