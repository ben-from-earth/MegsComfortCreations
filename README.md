# Meg's Comfort Creations

A Vite + React tool for the Etsy Shop Meg's Comfort Creations
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
2. Generate a single PNG for cricut cutting
3. Stores media information (specifically for books) for future workflows

## Prerequisites

1. Node.js - 18 or later
2. npm - 8 or later

## Installation

1. Clone to repository
    ```bash
    git clone https://github.com/ben-from-earth/MegsComfortCreations.git
    cd MegsComfortCreations
    ```
2. Install dependencies: The code is broken into client and server folders, so to install all dependencies:
    1. (in main directory)
        ```bash
        npm install && cd client && npm install && cd ../server && npm install
        ```

## Usage

To run the project, use the following command in the root:

```bash
npm start
```

## Project Structure

```text
.
├── client
│   ├── src
│   ├── index.html
│   ├── jsconfig.json
│   ├── package-lock.json
│   ├── package.json
│   └── vite.config.js
├── public
├── server
│   ├── __test__
│   ├── database
│   ├── documentation
│   ├── helpers
│   ├── models
│   ├── routes
│   ├── schemas
│   ├── MegsComfortCreations.sql
│   ├── MegsComfortCreations_test.sql
│   ├── app.js
│   ├── package-lock.json
│   ├── package.json
│   └── server.js
├── README.md
├── eslint.config.js
├── package-lock.json
└── package.json

```

## Environment Variables Required

Please see .env.example files in both /client and /server for full required file. Description of items is shown below.

1. client/
    1. VITE_SERVER_DOMAIN - The code is set up to run on localhost:3001, but this can be whatever you want. Just make sure that the client folder has a .env for this variable
2. root
    1. This tool requires a profile related to Google's search API. Follow Google documentation for setup, but access to the Search API requires the following:
        1. GOOGLE_SEARCH_API_KEY = long string given once profile is created
        2. GOOGLE_SEARCH_CX = long string given once profile is created
        3. email me at address below for these if just testing/checking out the app
    2. This tool is build on PostgreSQL for database functionality you will need:
        1. PG_USERNAME - PostrgreSQL username
        2. PG_PASSWORD - PostrgreSQL password
        3. DB_PORT - whichever port psql database is hosted on your machine (typically 5432)
        4. Create the main (/server/MegsComfortCreations.sql) and test (/server/MegsComfortCreations_test.sql) databases in the server directory: `bash psql < [file.sql]`

## Route Documentation

Route documentation can be found by hitting localhost:3001/docs, and the write up is found at server/documentation

## Contributing

Questions, ideas, or interested in collaborating? Email me at benknox480@gmail.com.

## Further Study

1. Build out other pages (Home, Shop, Meg's Recs, Newsletter, etc.) to turn into full shop website
2. Login/out and auth
3. Handling of multiple images per database item and using Google Cloud Storage for actual images instead of just holding urls
4. Update project to NextJS
5. More clear indication of where errors exist in MediaCollector Collected Cover Blocks
    1. i.e. red background if missing images, title, etc.
