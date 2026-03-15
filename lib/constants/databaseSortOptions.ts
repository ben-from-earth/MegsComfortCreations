export const DATABASE_SORT_OPTIONS = [
  'title',
  'author',
  'pageCount',
  'pubYear',
] as const;

export type DatabaseSortOption = (typeof DATABASE_SORT_OPTIONS)[number];

export const BOOK_SORT_OPTIONS: ReadonlyArray<DatabaseSortOption> = [
  'title',
  'author',
  'pageCount',
  'pubYear',
];

export const NON_BOOK_SORT_OPTIONS: ReadonlyArray<DatabaseSortOption> = ['title'];

export const BOOK_SORT_SELECT_OPTIONS: ReadonlyArray<{
  label: string;
  value: DatabaseSortOption;
}> = [
  { label: 'Title', value: 'title' },
  { label: 'Author', value: 'author' },
  { label: 'Page Count', value: 'pageCount' },
  { label: 'Pub. Year', value: 'pubYear' },
];

export const NON_BOOK_SORT_SELECT_OPTIONS: ReadonlyArray<{
  label: string;
  value: DatabaseSortOption;
}> = [{ label: 'Title', value: 'title' }];
