import {
  pgTable,
  pgEnum,
  uuid,
  text,
  integer,
  timestamp,
  boolean,
  index,
} from 'drizzle-orm/pg-core';
import { uniqueIndex } from 'drizzle-orm/pg-core';
import { relations } from 'drizzle-orm';
import { OTHER_MEDIA_TYPES } from 'lib/constants/media-types';

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

export const otherMediaTypeEnum = pgEnum('other_media_type', OTHER_MEDIA_TYPES);

// ---------- other_media ----------
export const otherMedia = pgTable(
  'other_media',
  {
    id: uuid('id').defaultRandom().primaryKey(),
    mediaType: otherMediaTypeEnum('media_type').notNull(),
    title: text('title').notNull(),
    spineColor: text('spine_color').notNull(),
  },
  (table) => [
    uniqueIndex('other_media_media_type_title_unique').on(
      table.mediaType,
      table.title,
    ),
    index('other_media_media_type_idx').on(table.mediaType),
    index('other_media_title_idx').on(table.title),
  ],
);

// ---------- media_image_items ----------
export const mediaImageItems = pgTable(
  'media_image_items',
  {
    id: uuid('id').defaultRandom().primaryKey(),
    bookId: uuid('book_id').references(() => books.id, { onDelete: 'cascade' }),
    otherMediaId: uuid('other_media_id').references(() => otherMedia.id, {
      onDelete: 'cascade',
    }),
    path: text('path').notNull(),
    sourceUrl: text('source_url'),
    mimeType: text('mime_type'),
    sizeBytes: integer('size_bytes'),
    sortOrder: integer('sort_order').notNull().default(0),
    isDefault: boolean('is_default').notNull().default(false),
    spineColor: text('spine_color').notNull().default('#ffffff'),
    createdAt: timestamp('created_at').defaultNow().notNull(),
    updatedAt: timestamp('updated_at')
      .defaultNow()
      .$onUpdate(() => new Date())
      .notNull(),
  },
  (table) => [
    index('media_image_items_book_id_idx').on(table.bookId),
    index('media_image_items_other_media_id_idx').on(table.otherMediaId),
    index('media_image_items_book_sort_order_idx').on(
      table.bookId,
      table.sortOrder,
    ),
    index('media_image_items_other_media_sort_order_idx').on(
      table.otherMediaId,
      table.sortOrder,
    ),
    index('media_image_items_book_is_default_idx').on(table.bookId, table.isDefault),
    index('media_image_items_other_media_is_default_idx').on(
      table.otherMediaId,
      table.isDefault,
    ),
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
