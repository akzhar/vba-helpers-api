# VBA helpers API

[API home page](https://vba-helpers-api.herokuapp.com/)

## API routes

- [GET /api/helpers](https://vba-helpers-api.herokuapp.com/api/helpers) - получение хелперов
- [GET /api/helpers/search-by-category/:keyword](https://vba-helpers-api.herokuapp.com/api/helpers/search-by-category/числа) - поиск хелперов по категории
- [GET /api/helpers/search-by-name/:keyword](https://vba-helpers-api.herokuapp.com/api/helpers/search-by-name/getlastrow) - поиск хелперов по имени
- [GET /api/helpers/search-by-title/:keyword](https://vba-helpers-api.herokuapp.com/api/helpers/search-by-title/получить%20индекс) - поиск хелперов по заголовку
- [GET /api/helpers/search-by-keyword/:keyword](https://vba-helpers-api.herokuapp.com/api/helpers/search-by-keyword/sort%20array) - поиск хелперов по ключевым словам (фразе)
- [GET /api/keywords/search-by-category/:keyword](https://vba-helpers-api.herokuapp.com/api/keywords/search-by-category/конвертация) - получение списка ключевых слов (фраз) по категории
- [GET /api/categories](https://vba-helpers-api.herokuapp.com/api/categories) - получение списка категорий
- [GET /api/categories/search/:keyword](https://vba-helpers-api.herokuapp.com/api/categories/search/строки) - поиск категории по имени

## 1. Установка зависимостей
Находясь в корневой папке проекта, выполните команду:
`npm i`

## 2. Запуск приложения в dev режиме
Находясь в корневой папке проекта, выполните команду:
`npm run start`

## 3. Описание функционала
- API раздает JSON данные для фронтенда