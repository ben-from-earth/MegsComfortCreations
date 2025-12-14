DROP DATABASE IF EXISTS megscomfortcreations_test;

CREATE DATABASE megscomfortcreations_test;

\c megscomfortcreations_test

CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

CREATE TABLE users (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    email TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    first_name TEXT NOT NULL,
    last_name TEXT NOT NULL 
);

CREATE TABLE books (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    title TEXT NOT NULL,
    author TEXT NOT NULL,
    pageCount INTEGER,
    pubYear INTEGER,
    spineColor TEXT NOT NULL,
    imageUrls TEXT[] NOT NULL,
    UNIQUE (title, author)
);

CREATE TABLE movies (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    title TEXT NOT NULL UNIQUE,
    spineColor TEXT NOT NULL,
    imageUrls TEXT[]
);

CREATE TABLE video_games (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    title TEXT NOT NULL UNIQUE,
    spineColor TEXT NOT NULL,
    imageUrls TEXT[]
);

CREATE TABLE albums (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    title TEXT NOT NULL UNIQUE,
    spineColor TEXT NOT NULL,
    imageUrls TEXT[]
);

CREATE TABLE genres (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    genre TEXT NOT NULL UNIQUE
);

CREATE TABLE genres_books (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    book_id  UUID NOT NULL,
    genre_id UUID NOT NULL,
    FOREIGN KEY (book_id) REFERENCES books(id)  ON DELETE CASCADE,
    FOREIGN KEY (genre_id) REFERENCES genres(id) ON DELETE CASCADE
);

INSERT INTO genres (
    genre
) VALUES 
    ('Children''s Fiction'),
    ('Middle-Grade Fiction'),
    ('Young Adult Fiction'),
    ('New Adult Fiction'),
    ('Romance'),
    ('Contemporary Fiction'),
    ('Spicy Romance (18+)'),
    ('LGBTQ'),
    ('Romantasy'),
    ('Fantasy'),
    ('Historical Fiction'),
    ('Mystery'),
    ('Thriller'),
    ('Horror'),
    ('Science Fiction'),
    ('Classic Literature'),
    ('Memoir'),
    ('History'),
    ('Philosophy'),
    ('Anthology');