# VBA helpers API

## 1. Description
This repository stores [VBA](https://en.wikipedia.org/wiki/Visual_Basic_for_Applications) helpers as `.bas` files inside the [data folder](https://github.com/akzhar/vba-helpers-api/tree/main/data) and provide JSON API to search helpers from the [VBA helpers application](https://github.com/akzhar/vba-helpers).

`helper` - utillity procedure / function.

## 2. Links
- [VBA helpers API](https://vba-helpers-api.herokuapp.com)
- [VBA helpers application repository (frontend)](https://github.com/akzhar/vba-helpers)

## 3. API routes
- [GET /api/helpers](https://vba-helpers-api.herokuapp.com/api/helpers) - get helpers data
- [GET /api/helpers/search-by-title/**:keyword**](https://vba-helpers-api.herokuapp.com/api/helpers/search-by-title/получить%20индекс) - search helpers by title
- [GET /api/helpers/search-by-category/**:keyword**](https://vba-helpers-api.herokuapp.com/api/helpers/search-by-category/числа) - search helpers by category
- [GET /api/helpers/search-by-keyword/**:keyword**](https://vba-helpers-api.herokuapp.com/api/helpers/search-by-keyword/sort%20array) - search helpers by keywords
- [GET /api/helpers/search-by-name/**:keyword**](https://vba-helpers-api.herokuapp.com/api/helpers/search-by-name/getlastrow) - search helpers by name
- [GET /api/categories](https://vba-helpers-api.herokuapp.com/api/categories) - get helpers categories
- [GET /api/categories/search/**:keyword**](https://vba-helpers-api.herokuapp.com/api/categories/search/строки) - search helpers category by its name
- [GET /api/categories/search-by-keyword/**:keyword**](https://vba-helpers-api.herokuapp.com/api/categories/search-by-keyword/конвертация) - search helpers category by keywords

## 4. Install dependencies
`git clone repo_url` → `cd ./repo-folder` → `npm install`

## 5. Build and run the app
`cd ./repo-folder` → `npm run start`