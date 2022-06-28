'use strict';

// ф-ция разбивает данные (json) по роутам
function getRoutedData(json) {
  
  let [ categories, keywordsByCategory ] = json.reduce(([ categories, keywordsByCategory ], item) => {
    
    const { category: itemCategories, _keywords: itemKeywords } = item;

    const keywords = itemKeywords.split('\n');

    itemCategories.forEach(category => {

      const categoryExists = Boolean(Object.hasOwnProperty.call(keywordsByCategory, category));

      if (categoryExists) {
        keywordsByCategory[category] = keywordsByCategory[category].concat(keywords);
      } else {
        keywordsByCategory[category] = keywords;
      }

      keywordsByCategory[category] = [...new Set(keywordsByCategory[category])];

    });

    return [ categories, keywordsByCategory ];

  }, [ {}, {} ]);

  categories = Object.keys(keywordsByCategory);
  categories = categories.map((item, i) => ({ id: i, category: item }));
  keywordsByCategory = Object.keys(keywordsByCategory).map((key, i) => ( { id: i, category: key, keywords: keywordsByCategory[key] }));

  return { helpers: json, categories, keywordsByCategory };
}

module.exports = getRoutedData;