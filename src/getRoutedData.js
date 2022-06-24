'use strict';

// ф-ция разбивает данные (json) по роутам
function getRoutedData(json) {
  
  let [ categories, tagsByCategory ] = json.reduce(([ categories, tagsByCategory ], item) => {
    
    const { category: itemCategories, tags: itemTags } = item;

    itemCategories.forEach(item => {
      if (!Object.hasOwnProperty.call(tagsByCategory, item)) {
        tagsByCategory[item] = itemTags;
      } else {
        tagsByCategory[item].concat(itemTags);
      }
      tagsByCategory[item] = [...new Set(tagsByCategory[item])];
    });

    return [ categories, tagsByCategory ];

  }, [ {}, {} ]);

  categories = Object.keys(tagsByCategory);
  categories = categories.map((item, i) => ({ id: i, category: item }));
  tagsByCategory = Object.keys(tagsByCategory).map((key, i) => ( { id: i, category: key, tags: tagsByCategory[key] }));

  return { helpers: json, categories, tagsByCategory };
}

module.exports = getRoutedData;