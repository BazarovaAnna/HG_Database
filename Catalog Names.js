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

var MASTERWORK_SUFFIX = ' (masterwork)';

function resolveCatalogItem(name, catalog) {
  var key = itemKey(name);
  if (catalog[key]) return catalog[key];

  var variant = key.match(/^(.+?)\s*\(\s*(?:masterwork|искусно сделанн(?:ый|ая|ое|ые))\s*\)$/);
  if (!variant) return null;
  var base = catalog[variant[1]];
  if (!base) return null;

  var variantName = base.name + MASTERWORK_SUFFIX;
  var exact = catalog[itemKey(variantName)];
  if (exact) return exact;

  var cost = Number(base.cost);
  var surcharge = Number(base.masterworkSurcharge);
  // A zero surcharge marks categories priced individually in the catalog.
  var hasPrice = String(base.cost).trim() !== '' &&
    String(base.masterworkSurcharge ?? '').trim() !== '' &&
    Number.isFinite(cost) && cost >= 0 && Number.isFinite(surcharge) && surcharge > 0;
  return {
    name: variantName,
    aliases: splitCatalogNames(base.aliases || '').map(function(alias) {
      return alias + MASTERWORK_SUFFIX;
    }).join('|'),
    cost: hasPrice ? cost + surcharge : '',
    weight: base.weight,
    link: base.link,
    category: base.category,
  };
}

function canonicalItemName(name, catalog) {
  var item = resolveCatalogItem(name, catalog);
  return item ? item.name : name;
}

function inventoryItemKey(character, itemName) {
  return character.toLowerCase() + '|' + itemName.toLowerCase();
}
