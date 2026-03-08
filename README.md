This is a [Next.js](https://nextjs.org) project bootstrapped with [`create-next-app`](https://nextjs.org/docs/app/api-reference/cli/create-next-app).

# Meg's Comfort Creations

A NextJS + React tool for the Etsy Shop Meg's Comfort Creations
Visit the Shop: [Etsy Link](https://www.etsy.com/shop/megscomfortcreations)

This tool streamlines the tedious process of gathering product cover images and laying them out into a single PNG.
The PNG is then used with a Cricut to cut out allthe cover images for in-store products.
The goal is simple: reduce the time from image collection to cutout as much as possible and store information about the item itself for possible use later on.

## Table of Contents

1.  [Features](#features)
2.  [Prerequisites](#prerequisites)
3.  [Installation](#installation)
4.  [Usage](#usage)
5.  [Project Structure](#project-structure)
6.  [Environment Variables](#environment-variables-required)
7.  [Route Documentation](#route-documentation)
8.  [Contributing](#contributing)

## Features

1. Collects media cover images programatically
2. Generate a single PNG for Cricut cutting
3. Stores media information (like auther, page count, and publication date) for future workflows

## Prerequisites

1. Node.js - 18 or later
2. npm - 8 or later
3. psql - 14 or later

## Installation

1. Clone to repository
    ```bash
    git clone https://github.com/ben-from-earth/MegsComfortCreations.git
    cd MegsComfortCreations
    ```
2. Install dependencies
    1. (in main directory)
        ```bash
        npm install
        ```

## Usage

To run the project, use the following command in the root:

```bash
npm run dev
```

## Project Structure

```text
.
├── __tests__
├── app
│   ├── api
│   │   ├── database
│   │   │   ├── delete
│   │   │   ├── edit
│   │   │   ├── save
│   │   │   ├── search
│   │   │   ├── route.ts
│   │   ├── genres
│   │   │   ├── addlink
│   │   │   ├── getall
│   │   │   ├── getforbook
│   │   │   ├── nogenres
│   │   │   ├── unlink
│   │   │   ├── route.ts
│   │   ├── getonlinedata
│   │   │   ├── mediacovers
│   │   │   ├── openlibrary
│   │   ├── png
│   │   │   ├── create
│   │   ├── profile
│   ├── components
│   ├── docs
│   ├── mediacollector
│   ├── profile
│   ├── showdatabase
│   ├── globals.css
│   ├── layout.tsx
│   ├── page.tsx
│   ├── Providers.tsx
├── public
├── lib
│   ├── context
│   ├── database
|   │   ├── models
|   │   ├── schemas
|   │   ├── databaseSeed
|   │   ├── config.ts
|   │   ├── db.ts
│   ├── helpers
│   ├── interfaces
│   ├── state
|   │   ├── slices
|   │   ├── store.ts
├── README.md
├── .env
├── eslint.config.mjs
├── jest.config.ts
├── next.config.ts
├── tsconfig.ts
├── postcss.config.mjs
├── package-lock.json
└── package.json

```

## Environment Variables Required

Please see .env.example files in both /client and /server for full required file. Description of items is shown below.

1. SERVER_BASE_URL - The code is set up to run on localhost:3000 in dev, but this can be whatever you want.
2. This tool requires a profile related to Google's search API. Follow Google documentation for setup, but access to the Search API requires the following:
    1. GOOGLE_SEARCH_API_KEY = long string given once profile is created
    2. GOOGLE_SEARCH_CX = long string given once profile is created
    3. email me at address below for more information
3. This tool is build on PostgreSQL for database functionality you will need:
    1. PG_USERNAME - PostrgreSQL username
    2. PG_PASSWORD - PostrgreSQL password
    3. DB_PORT - whichever port psql database is hosted on your machine (typically 5432)
    4. Create the main (/lib/database/databaseSeed/MegsComfortCreations.sql) and test (/lib/database/databaseSeed/MegsComfortCreations_test.sql) databases: `bash psql < [filename]`

## Neon Branch Safety (Dev vs Prod)

When using Neon branches, keep DB targets explicit:

1. Local app runtime URL in `.env.local`
2. Dev migration URL in `.env.development.local`
3. Prod migration URL in `.env.production.local`

Commands:

- Verify target DB and row counts:
  - `npm run db:check:local`
  - `npm run db:check:dev`
  - `npm run db:check:prod`
- Apply migrations intentionally:
  - `npm run db:migrate:dev`
  - `npm run db:migrate:prod`
- Full release sequence checklist:
  - `ai-assistance/RELEASE_CHECKLIST.md`

## Route Documentation

Route documentation can be found by hitting localhost:3000/docs, and the write up is found at /docs

## Contributing

Questions, ideas, or interested in collaborating? Email me at benknox480@gmail.com.

## Further Study

1. Build out other pages (Home, Shop, Meg's Recs, Newsletter, etc.) to turn into full shop website
2. Handling of multiple images per database item and using Google Cloud Storage for actual images instead of just holding urls
3. More clear indication of where errors exist in MediaCollector Collected Cover Blocks
    1. i.e. red background if missing images, title, etc.
