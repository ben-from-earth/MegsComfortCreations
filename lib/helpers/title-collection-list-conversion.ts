export interface titleOutputObj {
  title: string;
  author?: string;
}

export default function titleCollectionListConversion(
  list: string,
): titleOutputObj[] {
  const seperateItems = list
    .split(',')
    .map((i) => i.trim())
    .filter((i) => i !== '');

  const titleObjs = seperateItems.map((i) => {
    const titleInfo = i.split('/').map((each) => each.trim());
    const title = titleInfo[0];
    const author = titleInfo[1];
    return { title, author };
  });
  return titleObjs;
}
