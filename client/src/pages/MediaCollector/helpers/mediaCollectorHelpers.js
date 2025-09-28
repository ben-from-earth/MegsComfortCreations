const updateQueryCount = () => {
  const today = new Date().toISOString().split('T')[0];
  const storedDate = localStorage.getItem('lastQueryDate');

  if (storedDate !== today) {
    localStorage.setItem('queryCount', '0');
    localStorage.setItem('lastQueryDate', today);
  }

  let qCount = Number(localStorage.getItem('queryCount'));
  qCount++;
  localStorage.setItem('queryCount', `${qCount}`);
};

//updates to this function must be translated to serverside as well
const titleRearrange = (title) => {
  if (title.toLowerCase().startsWith('the ')) {
    const newTitle = title.slice(4) + ', The';
    return newTitle;
  }

  if (title.toLowerCase().startsWith('a ')) {
    const newTitle = title.slice(2) + ', A';
    return newTitle;
  }

  if (title.toLowerCase().startsWith('an ')) {
    const newTitle = title.slice(3) + ', An';
    return newTitle;
  }

  if (title.toLowerCase().endsWith(', the')) {
    const newTitle = 'The ' + title.slice(0, title.length - 5);
    return newTitle;
  }

  if (title.toLowerCase().endsWith(', a')) {
    const newTitle = 'A ' + title.slice(0, title.length - 3);
    return newTitle;
  }

  if (title.toLowerCase().endsWith(', an')) {
    const newTitle = 'An ' + title.slice(0, title.length - 4);
    return newTitle;
  }
  return title;
};

export { updateQueryCount, titleRearrange };
