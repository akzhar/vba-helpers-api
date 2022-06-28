# VBA helpers API

- [GET /helpers](https://vba-helpers-api.herokuapp.com/api/helpers) - получение хелперов
- [GET /helpers/search-by/category/:keyword](https://vba-helpers-api.herokuapp.com/api/helpers/search-by/category/числа) - поиск хелперов по категории
- [GET /helpers/search-by/title/:keyword](https://vba-helpers-api.herokuapp.com/api/helpers/search-by/title/получить%20индекс) - поиск хелперов по заголовку
- [GET /helpers/search-by/keyword/:keyword](https://vba-helpers-api.herokuapp.com/api/helpers/search-by/keyword/sort%20array) - поиск хелперов по ключевым словам (фразе)
- [GET /keywords/search-by/category/:keyword](https://vba-helpers-api.herokuapp.com/api/keywords/search-by/category/конвертация) - получение списка ключевых слов (фраз) по категории
- [GET /categories](https://vba-helpers-api.herokuapp.com/api/categories) - получение списка категорий
- [GET /categories/search/:keyword](https://vba-helpers-api.herokuapp.com/api/categories/search/строки) - поиск категории по имени

## 1. Установка зависимостей
Находясь в корневой папке проекта, выполните команду:
`npm i`

## 2. Запуск приложения в dev режиме
Находясь в корневой папке проекта, выполните команду:
`npm run start`

## 3. Описание функционала
- API раздает JSON данные для фронтенда