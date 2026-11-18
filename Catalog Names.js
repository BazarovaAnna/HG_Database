// The first name is the stable name used in transactions and exports.
function splitCatalogNames(value) {
  return String(value).split('|').map(function(name) {
    return name.trim();
  }).filter(Boolean);
}

function indexCatalogNames(index, names, entry, keyFor, sheetName, rowNumber) {
  names.forEach(function(name) {
    var key = keyFor(name);
    if (index[key] && index[key] !== entry) {
      throw new Error(sheetName + ' row ' + rowNumber + ': "' + name +
                      '" is already used for "' + index[key].name + '"');
    }
    index[key] = entry;
  });
}

function itemKey(name) {
  return String(name).trim().toLowerCase().replace(/ё/g, 'е').replace(/\s+/g, ' ');
}

function canonicalItemName(name, catalog) {
  var item = catalog[itemKey(name)];
  return item ? item.name : name;
}

function inventoryItemKey(character, itemName) {
  return character.toLowerCase() + '|' + itemName.toLowerCase();
}
