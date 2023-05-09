# VBA helpers API

## 1. What is it
This is a JSON API for [VBA helpers site](https://vbahelpers.ru).

And also a helpers storage.
## 2. Where are the helpers stored
[Excel file](https://github.com/akzhar/vba-helpers-api/tree/main/data) stores all the helper's data and [excel-to-json](https://github.com/akzhar/excel-to-json) converts it into JSON file.

All the [VBA](https://en.wikipedia.org/wiki/Visual_Basic_for_Applications) helpers are stored as plain text files in [./data/code](https://github.com/akzhar/vba-helpers-api/tree/main/data/code) folder.

## 3. Links
- [API (home page)](https://vbahelpers.ru:3001)
- [Frontend repository](https://github.com/akzhar/vba-helpers)

## 4. API routes
- [GET /api/helpers](https://vbahelpers.ru:3001/api/helpers) - get helpers data
- [GET /api/helpers/search-by-title/**:keyword**](https://vbahelpers.ru:3001/api/helpers/search-by-title/get%20index) - search helpers by title
- [GET /api/helpers/search-by-category/**:keyword**](https://vbahelpers.ru:3001/api/helpers/search-by-category/http) - search helpers by category
- [GET /api/helpers/search-by-keyword/**:keyword**](https://vbahelpers.ru:3001/api/helpers/search-by-keyword/sort%20array) - search helpers by keywords
- [GET /api/helpers/search-by-name/**:keyword**](https://vbahelpers.ru:3001/api/helpers/search-by-name/getlastrow) - search helpers by name
- [GET /api/categories](https://vbahelpers.ru:3001/api/categories) - get helpers categories
- [GET /api/categories/search/**:keyword**](https://vbahelpers.ru:3001/api/categories/search/text) - search helpers category by its name
- [GET /api/categories/search-by-keyword/**:keyword**](https://vbahelpers.ru:3001/api/categories/search-by-keyword/check%20if) - search helpers category by keywords

## 5. Install dependencies
`git clone repo_url` → `cd ./repo-folder` → `npm ci`

## 6. Build and run the app
`cd ./repo-folder` → `npm run start`