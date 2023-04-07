'use strict';

function compareFn(a, b) {
  if (a.name < b.name) { return -1; }
  if (a.name > b.name) { return 1; }
  return 0;
}

// ф-ция разбивает данные (json) по роутам
function getRoutedData(json) {

  json.sort(compareFn);

  let categoriesData = json.reduce((acc, item) => {
    
    const { category: itemCategories, _keywords } = item;

    const itemKeywords = _keywords.split('\n');

    itemCategories.forEach(category => {
      const categoryExists = Boolean(Object.hasOwnProperty.call(acc, category));

      const keywords = categoryExists ? acc[category].keywords.concat(itemKeywords) : itemKeywords;
      const helpersCount = categoryExists ? acc[category].helpersCount + 1 : 1;

      acc[category] = { keywords: [...new Set(keywords)], helpersCount };
    });

    return acc;

  }, {});

  const categories = Object.keys(categoriesData).sort().map((category, i) => {
    const { helpersCount, keywords } = categoriesData[category];
    return { id: i, category, helpersCount, keywords };
  });

  return { helpers: json, categories };
}

module.exports = getRoutedData;