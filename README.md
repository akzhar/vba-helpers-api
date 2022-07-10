# vba helpers api

## 1. Описание
`vba helpers api` - это API с данными по [VBA](https://ru.wikipedia.org/wiki/Visual_Basic_for_Applications) хелперам для [фронтенда](https://github.com/akzhar/vba-helpers).

`helper` - вспомогательная процедура / функция.

## 2. Ссылки
- [Опубликованное приложение](https://vba-helpers-api.herokuapp.com)
- [Фронтенд репозиторий](https://github.com/akzhar/vba-helpers)

## 3. Маршруты
- [GET /api/helpers](https://vba-helpers-api.herokuapp.com/api/helpers) - получение хелперов
- [GET /api/helpers/search-by-title/**:keyword**](https://vba-helpers-api.herokuapp.com/api/helpers/search-by-title/получить%20индекс) - поиск хелперов по заголовку
- [GET /api/helpers/search-by-category/**:keyword**](https://vba-helpers-api.herokuapp.com/api/helpers/search-by-category/числа) - поиск хелперов по категории
- [GET /api/helpers/search-by-keyword/**:keyword**](https://vba-helpers-api.herokuapp.com/api/helpers/search-by-keyword/sort%20array) - поиск хелперов по ключевым словам (фразе)
- [GET /api/helpers/search-by-name/**:keyword**](https://vba-helpers-api.herokuapp.com/api/helpers/search-by-name/getlastrow) - поиск хелперов по имени
- [GET /api/categories](https://vba-helpers-api.herokuapp.com/api/categories) - получение списка категорий
- [GET /api/categories/search/**:keyword**](https://vba-helpers-api.herokuapp.com/api/categories/search/строки) - поиск категории по имени
- [GET /api/categories/search-by-keyword/**:keyword**](https://vba-helpers-api.herokuapp.com/api/categories/search-by-keyword/конвертация) - поиск категорий по ключевым словам (фразе)

## 4. Установка зависимостей
Находясь в корневой папке проекта, выполните команду: `npm i`

## 5. Запуск приложения в dev режиме
Находясь в корневой папке проекта, выполните команду: `npm run start`