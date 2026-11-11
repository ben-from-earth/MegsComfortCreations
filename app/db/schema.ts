import {
  pgTable,
  pgEnum,
  uuid,
  text,
  integer,
  timestamp,
  boolean,
  date,
  index,
  uniqueIndex,
  check,
} from 'drizzle-orm/pg-core';
import { relations, sql } from 'drizzle-orm';
import { MOVIE_FORMATS } from 'lib/constants/movie-formats';

// ---------- users ----------
export const users = pgTable('users', {
  id: uuid('id').defaultRandom().primaryKey(),
  email: text('email').notNull().unique(),
  passwordHash: text('password_hash').notNull(),
  firstName: text('first_name').notNull(),
  lastName: text('last_name').notNull(),
});

// ---------- orders ----------
export const orders = pgTable('orders', {
  id: uuid('id').defaultRandom().primaryKey(),
  customerId: uuid('customer_id')
    .notNull()
    .references(() => customers.id, { onDelete: 'cascade' }),
  orderNumber: text('order_number').notNull().unique(),
  orderDate: timestamp('order_date').defaultNow().notNull(),
  totalAmount: integer('total_amount').notNull(),
  pngId: uuid('png_id').references(() => pngs.id),
});

// ---------- orders_books ----------
export const ordersBooks = pgTable('orders_books', {
  id: uuid('id').defaultRandom().primaryKey(),
  orderId: uuid('order_id')
    .notNull()
    .references(() => orders.id, { onDelete: 'cascade' }),
  bookId: uuid('book_id')
    .notNull()
    .references(() => books.id, { onDelete: 'cascade' }),
});

// ---------- pngs ----------
export const pngs = pgTable('pngs', {
  id: uuid('id').defaultRandom().primaryKey(),
  url: text('url').notNull(),
  description: text('description'),
});

// ---------- books ----------
export const books = pgTable(
  'books',
  {
    id: uuid('id').defaultRandom().primaryKey(),
    title: text('title').notNull(),
    author: text('author').notNull(),
    pageCount: integer('page_count'),
    pubYear: integer('pub_year'),
    spineColor: text('spine_color').notNull(),
  },
  (table) => [
    uniqueIndex('books_title_author_unique').on(table.title, table.author),
  ],
);

export const movieFormatEnum = pgEnum('movie_format', MOVIE_FORMATS);

// ---------- movies ----------
export const movies = pgTable(
  'movies',
  {
    id: uuid('id').defaultRandom().primaryKey(),
    title: text('title').notNull(),
    spineColor: text('spine_color').notNull(),
    movieFormat: movieFormatEnum('movie_format').notNull(),
  },
  (table) => [
    uniqueIndex('movies_title_format_unique').on(table.title, table.movieFormat),
  ],
);

// ---------- video_games ----------
export const videoGames = pgTable(
  'video_games',
  {
    id: uuid('id').defaultRandom().primaryKey(),
    title: text('title').notNull(),
    spineColor: text('spine_color').notNull(),
  },
  (table) => [uniqueIndex('video_games_title_unique').on(table.title)],
);

// ---------- albums ----------
export const albums = pgTable(
  'albums',
  {
    id: uuid('id').defaultRandom().primaryKey(),
    title: text('title').notNull(),
    artist: text('artist').notNull(),
    releaseDate: date('release_date'),
    spotifyAlbumUri: text('spotify_album_uri'),
    spineColor: text('spine_color').notNull(),
  },
  (table) => [
    uniqueIndex('albums_title_artist_unique').on(table.title, table.artist),
    uniqueIndex('albums_spotify_album_uri_unique')
      .on(table.spotifyAlbumUri)
      .where(sql`${table.spotifyAlbumUri} IS NOT NULL`),
  ],
);

// ---------- media_image_items ----------
export const mediaImageItems = pgTable(
  'media_image_items',
  {
    id: uuid('id').defaultRandom().primaryKey(),
    bookId: uuid('book_id').references(() => books.id, { onDelete: 'cascade' }),
    movieId: uuid('movie_id').references(() => movies.id, {
      onDelete: 'cascade',
    }),
    videoGameId: uuid('video_game_id').references(() => videoGames.id, {
      onDelete: 'cascade',
    }),
    albumId: uuid('album_id').references(() => albums.id, {
      onDelete: 'cascade',
    }),
    path: text('path').notNull(),
    sourceUrl: text('source_url'),
    mimeType: text('mime_type'),
    sizeBytes: integer('size_bytes'),
    sortOrder: integer('sort_order').notNull().default(0),
    isDefault: boolean('is_default').notNull().default(false),
    spineColor: text('spine_color').notNull().default('#ffffff'),
    codeColor: text('code_color'),
    createdAt: timestamp('created_at').defaultNow().notNull(),
    updatedAt: timestamp('updated_at')
      .defaultNow()
      .$onUpdate(() => new Date())
      .notNull(),
  },
  (table) => [
    check(
      'media_image_items_parent_check',
      sql`(
        (${table.bookId} IS NOT NULL)::int +
        (${table.movieId} IS NOT NULL)::int +
        (${table.videoGameId} IS NOT NULL)::int +
        (${table.albumId} IS NOT NULL)::int
      ) = 1`,
    ),
    check(
      'media_image_items_code_color_check',
      sql`${table.albumId} IS NOT NULL OR ${table.codeColor} IS NULL`,
    ),
    index('media_image_items_book_id_idx').on(table.bookId),
    index('media_image_items_movie_id_idx').on(table.movieId),
    index('media_image_items_video_game_id_idx').on(table.videoGameId),
    index('media_image_items_album_id_idx').on(table.albumId),
    index('media_image_items_book_sort_order_idx').on(
      table.bookId,
      table.sortOrder,
    ),
    index('media_image_items_movie_sort_order_idx').on(
      table.movieId,
      table.sortOrder,
    ),
    index('media_image_items_video_game_sort_order_idx').on(
      table.videoGameId,
      table.sortOrder,
    ),
    index('media_image_items_album_sort_order_idx').on(
      table.albumId,
      table.sortOrder,
    ),
    index('media_image_items_book_is_default_idx').on(
      table.bookId,
      table.isDefault,
    ),
    index('media_image_items_movie_is_default_idx').on(
      table.movieId,
      table.isDefault,
    ),
    index('media_image_items_video_game_is_default_idx').on(
      table.videoGameId,
      table.isDefault,
    ),
    index('media_image_items_album_is_default_idx').on(
      table.albumId,
      table.isDefault,
    ),
    uniqueIndex('media_image_items_book_default_unique')
      .on(table.bookId)
      .where(sql`${table.bookId} IS NOT NULL AND ${table.isDefault} = true`),
    uniqueIndex('media_image_items_movie_default_unique')
      .on(table.movieId)
      .where(sql`${table.movieId} IS NOT NULL AND ${table.isDefault} = true`),
    uniqueIndex('media_image_items_video_game_default_unique')
      .on(table.videoGameId)
      .where(
        sql`${table.videoGameId} IS NOT NULL AND ${table.isDefault} = true`,
      ),
    uniqueIndex('media_image_items_album_default_unique')
      .on(table.albumId)
      .where(sql`${table.albumId} IS NOT NULL AND ${table.isDefault} = true`),
  ],
);

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

// ---------- customers ----------
export const customers = pgTable('customers', {
  id: uuid('id').defaultRandom().primaryKey(),
  firstName: text('first_name').notNull(),
  lastName: text('last_name').notNull(),
  addressLine1: text('address_line_1').notNull(),
  addressLine2: text('address_line_2'),
  city: text('city').notNull(),
  state: text('state').notNull(),
  postalCode: text('postal_code').notNull(),
  country: text('country').notNull(),
  phone: text('phone'),
  createdAt: timestamp('created_at').defaultNow().notNull(),
  updatedAt: timestamp('updated_at')
    .defaultNow()
    .$onUpdate(() => new Date())
    .notNull(),
});

// ---------- customers_users ----------
export const customersUsers = pgTable(
  'customers_users',
  {
    id: uuid('id').defaultRandom().primaryKey(),
    customerId: uuid('customer_id')
      .notNull()
      .references(() => customers.id, { onDelete: 'cascade' }),
    userId: uuid('user_id')
      .notNull()
      .references(() => users.id, { onDelete: 'cascade' }),
    createdAt: timestamp('created_at').defaultNow().notNull(),
  },
  (table) => [
    uniqueIndex('customers_users_customer_user_unique').on(
      table.customerId,
      table.userId,
    ),
  ],
);

// ---------- google api query usage ----------
export const googleApiQueryUsage = pgTable('google_api_query_usage', {
  id: uuid('id').defaultRandom().primaryKey(),
  date: text('date').notNull().unique(),
  queryCount: integer('query_count').notNull().default(0),
});

// ---------- better-auth ----------

export const user = pgTable('user', {
  id: text('id').primaryKey(),
  name: text('name').notNull(),
  email: text('email').notNull().unique(),
  emailVerified: boolean('email_verified').default(false).notNull(),
  image: text('image'),
  createdAt: timestamp('created_at').defaultNow().notNull(),
  updatedAt: timestamp('updated_at')
    .defaultNow()
    .$onUpdate(() => /* @__PURE__ */ new Date())
    .notNull(),
  role: text('role'),
  banned: boolean('banned').default(false),
  banReason: text('ban_reason'),
  banExpires: timestamp('ban_expires'),
});

export const session = pgTable(
  'session',
  {
    id: text('id').primaryKey(),
    expiresAt: timestamp('expires_at').notNull(),
    token: text('token').notNull().unique(),
    createdAt: timestamp('created_at').defaultNow().notNull(),
    updatedAt: timestamp('updated_at')
      .$onUpdate(() => /* @__PURE__ */ new Date())
      .notNull(),
    ipAddress: text('ip_address'),
    userAgent: text('user_agent'),
    userId: text('user_id')
      .notNull()
      .references(() => user.id, { onDelete: 'cascade' }),
    impersonatedBy: text('impersonated_by'),
  },
  (table) => [index('session_userId_idx').on(table.userId)],
);

export const account = pgTable(
  'account',
  {
    id: text('id').primaryKey(),
    accountId: text('account_id').notNull(),
    providerId: text('provider_id').notNull(),
    userId: text('user_id')
      .notNull()
      .references(() => user.id, { onDelete: 'cascade' }),
    accessToken: text('access_token'),
    refreshToken: text('refresh_token'),
    idToken: text('id_token'),
    accessTokenExpiresAt: timestamp('access_token_expires_at'),
    refreshTokenExpiresAt: timestamp('refresh_token_expires_at'),
    scope: text('scope'),
    password: text('password'),
    createdAt: timestamp('created_at').defaultNow().notNull(),
    updatedAt: timestamp('updated_at')
      .$onUpdate(() => /* @__PURE__ */ new Date())
      .notNull(),
  },
  (table) => [index('account_userId_idx').on(table.userId)],
);

export const verification = pgTable(
  'verification',
  {
    id: text('id').primaryKey(),
    identifier: text('identifier').notNull(),
    value: text('value').notNull(),
    expiresAt: timestamp('expires_at').notNull(),
    createdAt: timestamp('created_at').defaultNow().notNull(),
    updatedAt: timestamp('updated_at')
      .defaultNow()
      .$onUpdate(() => /* @__PURE__ */ new Date())
      .notNull(),
  },
  (table) => [index('verification_identifier_idx').on(table.identifier)],
);

export const userRelations = relations(user, ({ many }) => ({
  sessions: many(session),
  accounts: many(account),
}));

export const sessionRelations = relations(session, ({ one }) => ({
  user: one(user, {
    fields: [session.userId],
    references: [user.id],
  }),
}));

export const accountRelations = relations(account, ({ one }) => ({
  user: one(user, {
    fields: [account.userId],
    references: [user.id],
  }),
}));
