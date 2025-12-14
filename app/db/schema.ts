import { pgTable, uuid, text, integer } from 'drizzle-orm/pg-core';
import { uniqueIndex } from 'drizzle-orm/pg-core';

// ---------- users ----------
export const users = pgTable('users', {
  id: uuid('id').defaultRandom().primaryKey(),
  email: text('email').notNull().unique(),
  passwordHash: text('password_hash').notNull(),
  firstName: text('first_name').notNull(),
  lastName: text('last_name').notNull(),
});

// ---------- books ----------
export const books = pgTable(
  'books',
  {
    id: uuid('id').defaultRandom().primaryKey(),
    title: text('title').notNull(),
    author: text('author').notNull(),
    pageCount: integer('pageCount'),
    pubYear: integer('pubYear'),
    spineColor: text('spineColor').notNull(),
    imageUrls: text('imageUrls').array().notNull(),
  },
  (table) => ({
    titleAuthorUnique: uniqueIndex('books_title_author_unique').on(
      table.title,
      table.author,
    ),
  }),
);

// ---------- movies ----------
export const movies = pgTable('movies', {
  id: uuid('id').defaultRandom().primaryKey(),
  title: text('title').notNull().unique(),
  spineColor: text('spineColor').notNull(),
  imageUrls: text('imageUrls').array().notNull(), // nullable in SQL
});

// ---------- video_games ----------
export const videoGames = pgTable('video_games', {
  id: uuid('id').defaultRandom().primaryKey(),
  title: text('title').notNull().unique(),
  spineColor: text('spineColor').notNull(),
  imageUrls: text('imageUrls').array().notNull(),
});

// ---------- albums ----------
export const albums = pgTable('albums', {
  id: uuid('id').defaultRandom().primaryKey(),
  title: text('title').notNull().unique(),
  spineColor: text('spineColor').notNull(),
  imageUrls: text('imageUrls').array().notNull(),
});

// ---------- genres ----------
export const genres = pgTable('genres', {
  id: uuid('id').defaultRandom().primaryKey(),
  genre: text('genre').notNull().unique(),
});

// ---------- genres_books ----------
export const genresBooks = pgTable('genres_books', {
  id: uuid('id').defaultRandom().primaryKey(),
  bookId: uuid('book_id')
    .notNull()
    .references(() => books.id, { onDelete: 'cascade' }),
  genreId: uuid('genre_id')
    .notNull()
    .references(() => genres.id, { onDelete: 'cascade' }),
});
