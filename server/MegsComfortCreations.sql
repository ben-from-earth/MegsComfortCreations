DROP DATABASE IF EXISTS megscomfortcreations;

CREATE DATABASE megscomfortcreations;

\c megscomfortcreations

CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

CREATE TABLE books (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    title TEXT NOT NULL,
    author TEXT NOT NULL,
    page_count INTEGER,
    pub_year INTEGER ,
    spine_color TEXT,
    image_urls TEXT[]
);

CREATE TABLE movies (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    title TEXT NOT NULL,
    image_urls TEXT[]
);

CREATE TABLE video_games (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    title TEXT NOT NULL,
    image_urls TEXT[]
);

CREATE TABLE albums (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    title TEXT NOT NULL,
    image_urls TEXT[]
);

CREATE TABLE genres (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    genre TEXT NOT NULL
);

CREATE TABLE genres_books (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    book_id TEXT,
    genre_id TEXT
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
    ('Memoirs'),
    ('History'),
    ('Philosophy'),
    ('Anthologies');