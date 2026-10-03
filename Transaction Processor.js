// ============================================================
// PF1e Character Database — Inventory & Game Session Processor
// ============================================================
//
// Reads column mappings from "Column matchups" sheet — no hardcoded
// column numbers. If you move columns in the forms, just update
// the matchups sheet and the script adapts.
//
// COLUMN MATCHUPS SHEET FORMAT:
//   Row 1: Section headers (Registration, LvlUP, Inventory(raw), Games)
//   Row 2: "Field", "Column", "Form Question" headers
//   Row 3+: field name | column number | form question text
//   Sections at columns: Registration=1, LvlUP=5, Inventory=9, Games=13
// ============================================================


// ============================================================
// DYNAMIC COLUMN LOADER
// ============================================================

/**
 * Reads the Column matchups sheet and returns an object with
 * column mappings for each section.
 * Returns: { registration: {field: col, ...}, lvlup: {...}, inventory: {...}, games: {...} }
 */
function loadColumnMappings(ss) {
  var sheet = ss.getSheetByName('Column matchups');
  if (!sheet) throw new Error('Sheet "Column matchups" not found!');

  var data = sheet.getDataRange().getValues();

  // Section start columns (0-indexed): Field col, Number col
  var sections = {
    registration: { fieldCol: 0, numCol: 1 },
    lvlup:        { fieldCol: 4, numCol: 5 },
    inventory:    { fieldCol: 8, numCol: 9 },
    games:        { fieldCol: 12, numCol: 13 },
  };

  var mappings = {};
  for (var key in sections) {
    mappings[key] = {};
    var fc = sections[key].fieldCol;
    var nc = sections[key].numCol;
    for (var i = 2; i < data.length; i++) { // skip header rows (0, 1)
      var field = String(data[i][fc] || '').trim();
      var colNum = Number(data[i][nc]) || 0;
      if (field && colNum > 0) {
        mappings[key][field] = colNum;
      }
    }
  }

  return mappings;
}


/**
 * Helper: read a value from a row by field name using mappings.
 */
function getField(row, mappings, field) {
  var col = mappings[field];
  if (!col) return '';
  var val = row[col - 1]; // 0-indexed
  return (val === undefined || val === null) ? '' : val;
}


// ============================================================
// TRANSACTIONS SHEET COLUMNS (these are ours — we control them)
// ============================================================
var TXN_HEADERS = ['timestamp', 'source', 'character', 'item_name',
                   'quantity', 'cost', 'weight', 'notes'];
var TXN = {
  timestamp: 1, source: 2, character: 3, itemName: 4,
  quantity: 5, cost: 6, weight: 7, notes: 8,
};


// ============================================================
// MAIN ENTRY POINTS
// ============================================================

function processGameSessions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gameSheet = ss.getSheetByName('Games');
  if (!gameSheet) { SpreadsheetApp.getUi().alert('Sheet "Games" not found!'); return; }

  var cols = loadColumnMappings(ss);
  var G = cols.games;

  var txnSheet = getOrCreateTransactionsSheet(ss);
  var processedSet = getTxnProcessedSet(txnSheet);
  var gameData = gameSheet.getDataRange().getValues();
  var catalog = loadItemCatalog(ss);

  var processed = 0;
  var skipped = 0;
  var errors = [];

  for (var i = 1; i < gameData.length; i++) {
    var row = gameData[i];
    var timestamp = String(getField(row, G, 'timestamp'));
    var characters = String(getField(row, G, 'characters') || '').trim();

    if (!characters) continue;

    // Games has no Approve column — process all non-empty rows
    var processKey = timestamp + '|game';
    if (processedSet[processKey]) { skipped++; continue; }

    try {
      var txns = parseGameForm(row, G, catalog);
      writeTxnRows(txnSheet, txns);
      processed++;
      processedSet[processKey] = true;
    } catch (e) {
      errors.push('Game row ' + (i + 1) + ': ' + e.message);
    }
  }

  var msg = 'Game Sessions Processed!\n\n' +
    '✅ Processed: ' + processed + '\n' +
    '⏭ Skipped (already done): ' + skipped + '\n';
  if (errors.length > 0) {
    msg += '❌ Errors:\n';
    for (var e = 0; e < errors.length; e++) msg += '  • ' + errors[e] + '\n';
  }
  SpreadsheetApp.getUi().alert(msg);
}


function initializeStartingGold() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var db = ss.getSheetByName('Database');
  var txnSheet = getOrCreateTransactionsSheet(ss);
  if (!db) { SpreadsheetApp.getUi().alert('Database not found!'); return; }

  var data = db.getDataRange().getValues();
  var headers = data[3]; // row 4

  // Find columns by key
  var col = {};
  for (var c = 0; c < headers.length; c++) {
    var key = String(headers[c]).trim();
    if (key) col[key] = c;
  }

  var inventory = buildInventoryMap(txnSheet);

  var created = 0;
  var skipped = 0;

  for (var i = 4; i < data.length; i++) {
    var charName = String(data[i][col['char']] || '').trim();
    if (!charName) continue;

    // Skip if character already has gold in Transactions
    var key = charName.toLowerCase() + '|gold';
    if (key in inventory) { skipped++; continue; }

    var startingGold = Number(data[i][col['starting_gold']]) || 0;
    if (startingGold === 0) { skipped++; continue; }

    writeTxnRows(txnSheet, [{
      timestamp: new Date().toISOString(),
      source: 'registration',
      character: charName,
      itemName: 'gold',
      quantity: startingGold,
      cost: 1,
      weight: '',
      notes: 'Starting gold',
    }]);
    created++;
  }

  SpreadsheetApp.getUi().alert(
    'Starting Gold Initialized!\n\n' +
    '✅ Created: ' + created + '\n' +
    '⏭ Skipped (already exists): ' + skipped
  );
}


function processInventory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName('Inventory(raw)');
  if (!invSheet) { SpreadsheetApp.getUi().alert('Sheet "Inventory(raw)" not found!'); return; }

  var cols = loadColumnMappings(ss);
  var I = cols.inventory;

  var txnSheet = getOrCreateTransactionsSheet(ss);
  var processedSet = getTxnProcessedSet(txnSheet);
  var invData = invSheet.getDataRange().getValues();

  var inventory = buildInventoryMap(txnSheet);
  var catalog = loadItemCatalog(ss);

  var processed = 0;
  var skipped = 0;
  var errors = [];
  var warnings = [];

  for (var i = 1; i < invData.length; i++) {
    var row = invData[i];
    var timestamp = String(getField(row, I, 'timestamp'));
    var charName = String(getField(row, I, 'charName') || '').trim();
    var approved = getField(row, I, 'approved');

    if (!isApproved(approved)) continue;
    if (!charName) continue;

    var processKey = timestamp + '|inventory|' + charName.toLowerCase();
    if (processedSet[processKey]) { skipped++; continue; }

    try {
      // parseInventoryForm updates the inventory map itself, txn by txn,
      // so an item received earlier in the same answer can be sold in it
      var result = parseInventoryForm(row, I, inventory, charName, catalog);
      writeTxnRows(txnSheet, result.txns);
      for (var w = 0; w < result.warnings.length; w++) {
        warnings.push(charName + ': ' + result.warnings[w]);
      }
      processed++;
      processedSet[processKey] = true;
    } catch (e) {
      errors.push(charName + ': ' + e.message);
    }
  }

  var msg = 'Inventory Processed!\n\n' +
    '✅ Processed: ' + processed + '\n' +
    '⏭ Skipped (already done): ' + skipped + '\n';
  if (warnings.length > 0) {
    msg += '⚠️ Warnings:\n';
    for (var w = 0; w < warnings.length; w++) msg += '  • ' + warnings[w] + '\n';
  }
  if (errors.length > 0) {
    msg += '❌ Errors:\n';
    for (var e = 0; e < errors.length; e++) msg += '  • ' + errors[e] + '\n';
  }
  SpreadsheetApp.getUi().alert(msg);
}


function rebuildInventory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var txnSheet = ss.getSheetByName('Transactions');
  if (!txnSheet) { SpreadsheetApp.getUi().alert('No Transactions sheet found!'); return; }

  buildInventorySheet(ss, txnSheet);
  SpreadsheetApp.getUi().alert('Inventory sheet rebuilt from Transactions.');
}


// ============================================================
// GAME FORM PARSER
// ============================================================

function parseGameForm(row, G, catalog) {
  var txns = [];
  var timestamp = String(getField(row, G, 'timestamp'));
  var gameName = String(getField(row, G, 'gameName') || '').trim();
  var characters = String(getField(row, G, 'characters') || '').trim();
  var partyExp = Number(getField(row, G, 'exp')) || 0;
  var partyGold = Number(getField(row, G, 'goldReward')) || 0;
  var goldDeltaStr = String(getField(row, G, 'goldDelta') || '').trim();
  var lootStr = String(getField(row, G, 'loot') || '').trim();

  var sessionLabel = gameName || 'Session';
  var charList = characters.split(',').map(function(c) { return c.trim(); }).filter(function(c) { return c; });

  // --- Party gold reward for each character ---
  if (partyGold !== 0) {
    for (var ci = 0; ci < charList.length; ci++) {
      txns.push({
        timestamp: timestamp, source: 'game', character: charList[ci],
        itemName: 'gold', quantity: partyGold, cost: 1, weight: '',
        notes: sessionLabel + ' reward',
      });
    }
  }

  // --- Per-character gold changes: "Aurelia: -10; Grok: +13" ---
  if (goldDeltaStr) {
    var goldParts = goldDeltaStr.split(';');
    for (var g = 0; g < goldParts.length; g++) {
      var colonIdx = goldParts[g].indexOf(':');
      if (colonIdx === -1) continue;
      var charName = goldParts[g].substring(0, colonIdx).trim();
      var amount = Number(goldParts[g].substring(colonIdx + 1).trim()) || 0;
      if (charName && amount !== 0) {
        txns.push({
          timestamp: timestamp, source: 'game', character: charName,
          itemName: 'gold', quantity: amount, cost: 1, weight: '',
          notes: sessionLabel + ' gold change',
        });
      }
    }
  }

  // --- Additional rewards: "Warthud: Scroll 1 30, potion 2 15; Grok: bark; Aurelia: -mirror" ---
  if (lootStr) {
    var charBlocks = lootStr.split(';');
    for (var b = 0; b < charBlocks.length; b++) {
      var colonIdx = charBlocks[b].indexOf(':');
      if (colonIdx === -1) continue;
      var charName = charBlocks[b].substring(0, colonIdx).trim();
      var itemsStr = charBlocks[b].substring(colonIdx + 1).trim();
      if (!charName || !itemsStr) continue;

      var items = itemsStr.split(',');
      for (var it = 0; it < items.length; it++) {
        var parsed = applyCatalog(parseItemFull(items[it].trim()), catalog);
        if (!parsed) continue;
        txns.push({
          timestamp: timestamp, source: 'game', character: charName,
          itemName: parsed.name, quantity: parsed.quantity,
          cost: parsed.cost, weight: parsed.weight,
          notes: sessionLabel + ' loot',
        });
      }
    }
  }

  return txns;
}


// ============================================================
// INVENTORY FORM PARSER
// ============================================================

function parseInventoryForm(row, I, inventory, charName, catalog) {
  var txns = [];
  var warnings = [];
  var timestamp = String(getField(row, I, 'timestamp'));
  var comment = String(getField(row, I, 'comment') || '').trim();

  // Every txn goes through here so the inventory map stays current:
  // an item got earlier in this answer can be sold later in it.
  function add(txn) {
    txns.push(txn);
    updateInventoryMap(inventory, txn);
  }
  function list(field) {
    var s = String(getField(row, I, field) || '').trim();
    return s ? s.split(',') : [];
  }

  // --- Items bought: cost and weight from the catalog if omitted ---
  var bought = list('itemsBought');
  for (var i = 0; i < bought.length; i++) {
    var parsed = applyCatalog(parseItemFull(bought[i].trim()), catalog);
    if (!parsed) continue;

    // An explicitly written 0 is a choice; an omitted price with no
    // catalog entry is probably a mistake — record it, but say so
    if (parsed.given < 2 && !(Number(parsed.cost) > 0)) {
      warnings.push('"' + parsed.name + '" bought for 0: no price stated and not in catalog');
    }

    var qty = Math.abs(parsed.quantity);
    add({
      timestamp: timestamp, source: 'inventory', character: charName,
      itemName: parsed.name, quantity: qty,
      cost: parsed.cost, weight: parsed.weight,
      notes: comment,
    });

    var totalCost = roundGold(parsed.cost * qty);
    if (totalCost > 0) {
      add({
        timestamp: timestamp, source: 'inventory', character: charName,
        itemName: 'gold', quantity: -totalCost, cost: 1, weight: '',
        notes: 'Bought ' + (qty > 1 ? qty + ' ' : '') + parsed.name,
      });
    }
  }

  // --- Items got: free, cost and weight from the catalog if omitted ---
  var got = list('itemsGot');
  for (var i = 0; i < got.length; i++) {
    var parsed = applyCatalog(parseItemFull(got[i].trim()), catalog);
    if (!parsed) continue;

    add({
      timestamp: timestamp, source: 'inventory', character: charName,
      itemName: parsed.name, quantity: Math.abs(parsed.quantity),
      cost: parsed.cost, weight: parsed.weight,
      notes: comment,
    });
  }

  // --- Items sold: only what the character has ---
  // "name qty price": price is what the character got per item.
  // Without a price: half the catalog price, else half the price it was
  // bought for, else nothing (with a warning).
  var sold = list('itemsSold');
  for (var i = 0; i < sold.length; i++) {
    var parsed = parseItemFull(sold[i].trim());
    if (!parsed) continue;

    var qty = Math.abs(parsed.quantity);
    var lookup = lookupItem(inventory, charName, parsed.name);
    if (!lookup.found || lookup.quantity < qty) {
      warnings.push('"' + parsed.name + '" x' + qty + ' not sold: character has ' +
                    (lookup.found ? lookup.quantity : 0));
      continue;
    }

    add({
      timestamp: timestamp, source: 'inventory', character: charName,
      itemName: parsed.name, quantity: -qty,
      cost: lookup.cost, weight: lookup.weight,
      notes: comment,
    });

    var price = null, how = '';
    var ref = catalog[parsed.name.toLowerCase()];
    if (parsed.given >= 2) {
      price = parsed.cost; how = 'stated price';
    } else if (ref && Number(ref.cost) > 0) {
      price = Number(ref.cost) / 2; how = 'half catalog price';
    } else if (Number(lookup.cost) > 0) {
      price = Number(lookup.cost) / 2; how = 'half purchase price';
    }

    if (price === null) {
      warnings.push('"' + parsed.name + '" sold for 0: no price stated, none in catalog or inventory');
      continue;
    }
    var sellGold = roundGold(price * qty);
    if (sellGold > 0) {
      add({
        timestamp: timestamp, source: 'inventory', character: charName,
        itemName: 'gold', quantity: sellGold, cost: 1, weight: '',
        notes: 'Sold ' + (qty > 1 ? qty + ' ' : '') + parsed.name + ' (' + how + ')',
      });
    }
  }

  // --- Items gifted: only what the character has ---
  var gifted = list('itemsGifted');
  for (var i = 0; i < gifted.length; i++) {
    var parsed = parseItemSimple(gifted[i].trim());
    if (!parsed) continue;

    var qty = Math.abs(parsed.quantity);
    var lookup = lookupItem(inventory, charName, parsed.name);
    if (!lookup.found || lookup.quantity < qty) {
      warnings.push('"' + parsed.name + '" x' + qty + ' not gifted: character has ' +
                    (lookup.found ? lookup.quantity : 0));
      continue;
    }

    add({
      timestamp: timestamp, source: 'inventory', character: charName,
      itemName: parsed.name, quantity: -qty,
      cost: lookup.cost, weight: lookup.weight,
      notes: comment,
    });
  }

  // --- Gold changes ---
  var goldChange = Number(getField(row, I, 'goldDelta')) || 0;
  if (goldChange !== 0) {
    add({
      timestamp: timestamp, source: 'inventory', character: charName,
      itemName: 'gold', quantity: goldChange, cost: 1, weight: '',
      notes: comment,
    });
  }

  return { txns: txns, warnings: warnings };
}

// Gold to the copper: 2.5 gp stays 2.5, float noise is cut off
function roundGold(x) {
  return Math.round(x * 100) / 100;
}


// ============================================================
// ITEM CATALOG
// ============================================================
// Sheet "Items catalog" is filled by hand: item_name | category | description |
// cost | weight | link. Items given by name in a form take cost and
// weight from here unless the form states them; the Inventory sheet
// takes the link from here unless one was entered by hand.

function loadItemCatalog(ss) {
  var catalog = {};
  var sheet = ss.getSheetByName('Items catalog');
  if (!sheet || sheet.getLastRow() <= 1) return catalog;

  var data = sheet.getDataRange().getValues();
  var hdr = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var c = {
    name: hdr.indexOf('item_name'), category: hdr.indexOf('category'),
    description: hdr.indexOf('description'), cost: hdr.indexOf('cost'),
    weight: hdr.indexOf('weight'), link: hdr.indexOf('link'),
  };
  if (c.name === -1) return catalog;

  for (var r = 1; r < data.length; r++) {
    var name = String(data[r][c.name]).trim();
    if (!name) continue;
    var row = data[r];
    var cell = function(k) { return c[k] === -1 ? '' : row[c[k]]; };
    catalog[name.toLowerCase()] = {
      name: name, category: cell('category'), description: cell('description'),
      cost: cell('cost'), weight: cell('weight'), link: String(cell('link') || '').trim(),
    };
  }
  return catalog;
}

// Fills what the form did not state: cost when fewer than 2 numbers
// were given, weight when fewer than 3. Stated values always win.
function applyCatalog(parsed, catalog) {
  if (!parsed || !catalog) return parsed;
  var ref = catalog[parsed.name.toLowerCase()];
  if (!ref) return parsed;
  if (parsed.given < 2 && Number(ref.cost) > 0) parsed.cost = Number(ref.cost);
  if (parsed.given < 3 && Number(ref.weight) > 0) parsed.weight = Number(ref.weight);
  return parsed;
}


// ============================================================
// ITEM PARSERS
// ============================================================

/**
 * Full format: "item name [qty] [cost] [weight]"
 * Numbers consumed from right. "-name" = negative qty.
 */
function parseItemFull(str) {
  if (!str) return null;
  str = str.trim();
  if (!str) return null;

  var negative = false;
  if (str.charAt(0) === '-') {
    negative = true;
    str = str.substring(1).trim();
  }

  var parts = str.split(/\s+/);
  if (parts.length === 0) return null;

  var numbers = [];
  while (parts.length > 0 && !isNaN(parts[parts.length - 1]) && parts[parts.length - 1] !== '') {
    numbers.unshift(Number(parts.pop()));
  }

  var name = parts.join(' ').trim();
  if (!name) return null;

  var quantity = 1, cost = 0, weight = '';

  if (numbers.length >= 3) {
    quantity = numbers[0];
    cost = numbers[1];
    weight = numbers[2] > 0 ? numbers[2] : '';
  } else if (numbers.length === 2) {
    quantity = numbers[0];
    cost = numbers[1];
  } else if (numbers.length === 1) {
    quantity = numbers[0];
  }

  if (negative) quantity = -Math.abs(quantity);

  // given: how many numbers were written, tells "omitted" from "0"
  return { name: name, quantity: quantity, cost: cost, weight: weight, given: numbers.length };
}

/**
 * Simple format: "item name [qty]"
 * Only quantity, no cost or weight.
 */
function parseItemSimple(str) {
  if (!str) return null;
  str = str.trim();
  if (!str) return null;

  var parts = str.split(/\s+/);
  if (parts.length === 0) return null;

  var quantity = 1;
  var lastToken = parts[parts.length - 1];
  if (!isNaN(lastToken) && lastToken !== '' && parts.length > 1) {
    quantity = Number(parts.pop());
  }

  var name = parts.join(' ').trim();
  if (!name) return null;

  return { name: name, quantity: quantity };
}


// ============================================================
// INVENTORY LOOKUP
// ============================================================

function buildInventoryMap(txnSheet) {
  var map = {};
  if (txnSheet.getLastRow() <= 1) return map;

  var data = txnSheet.getRange(2, 1, txnSheet.getLastRow() - 1, 8).getValues();
  for (var i = 0; i < data.length; i++) {
    var charName = String(data[i][TXN.character - 1]).trim();
    var itemName = String(data[i][TXN.itemName - 1]).trim();
    var qty = Number(data[i][TXN.quantity - 1]) || 0;
    var cost = data[i][TXN.cost - 1];
    var weight = data[i][TXN.weight - 1];

    if (!charName || !itemName) continue;

    var key = charName.toLowerCase() + '|' + itemName.toLowerCase();
    if (!(key in map)) {
      map[key] = { quantity: 0, cost: cost, weight: weight };
    }
    map[key].quantity += qty;
    if (cost && cost !== 'N/A' && Number(cost) > 0) map[key].cost = cost;
    if (weight && weight !== '' && Number(weight) > 0) map[key].weight = weight;
  }

  return map;
}

function updateInventoryMap(map, txn) {
  var key = txn.character.toLowerCase() + '|' + txn.itemName.toLowerCase();
  if (!(key in map)) {
    map[key] = { quantity: 0, cost: txn.cost, weight: txn.weight };
  }
  map[key].quantity += txn.quantity;
  if (txn.cost && txn.cost !== 'N/A' && Number(txn.cost) > 0) map[key].cost = txn.cost;
  if (txn.weight && txn.weight !== '' && Number(txn.weight) > 0) map[key].weight = txn.weight;
}

function lookupItem(map, charName, itemName) {
  var key = charName.toLowerCase() + '|' + itemName.toLowerCase();
  if (key in map && map[key].quantity > 0) {
    return { found: true, quantity: map[key].quantity, cost: map[key].cost || 0, weight: map[key].weight || '' };
  }
  return { found: false, quantity: 0, cost: 'N/A', weight: '' };
}


// ============================================================
// TRANSACTIONS SHEET
// ============================================================

function getOrCreateTransactionsSheet(ss) {
  var sheet = ss.getSheetByName('Transactions');
  if (!sheet) {
    sheet = ss.insertSheet('Transactions');
    sheet.getRange(1, 1, 1, TXN_HEADERS.length).setValues([TXN_HEADERS]);

    var headerRange = sheet.getRange(1, 1, 1, TXN_HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1B1B2F');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setHorizontalAlignment('center');

    sheet.setFrozenRows(1);
    sheet.setTabColor('#FF6600');
  }
  return sheet;
}

function getTxnProcessedSet(txnSheet) {
  var set = {};
  if (txnSheet.getLastRow() <= 1) return set;

  var data = txnSheet.getRange(2, 1, txnSheet.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < data.length; i++) {
    var ts = String(data[i][0]);
    var source = String(data[i][1]).trim().toLowerCase();
    var charName = String(data[i][2]).trim().toLowerCase();

    if (source === 'game') {
      set[ts + '|game'] = true;
    } else {
      set[ts + '|inventory|' + charName] = true;
    }
  }
  return set;
}

function writeTxnRows(txnSheet, txns) {
  if (txns.length === 0) return;

  var rows = [];
  for (var i = 0; i < txns.length; i++) {
    var t = txns[i];
    rows.push([t.timestamp, t.source, t.character, t.itemName,
               t.quantity, t.cost, t.weight, t.notes]);
  }

  var startRow = txnSheet.getLastRow() + 1;
  txnSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}


// ============================================================
// INVENTORY SHEET BUILDER
// ============================================================

// Columns of the Inventory sheet that come from Items catalog but may be
// overridden by hand. Order = order in the sheet, after "notes".
var INVENTORY_CATALOG_COLS = ['link', 'category'];

// Manual columns of the Inventory sheet, kept across rebuilds by
// character + item: { key: { notes, link, category } }.
// A catalog-backed value counts as manual unless it equals its cell note:
// values filled from Items catalog carry that same value as the note, so
// one the GM typed or changed differs from it and is kept, while an
// untouched catalog value is refilled from the current catalog.
function readInventoryManual(sheet) {
  var kept = {};
  if (sheet.getLastRow() <= 1) return kept;

  var data = sheet.getDataRange().getValues();
  var hdr = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var ch = hdr.indexOf('character'), item = hdr.indexOf('item_name'), notesCol = hdr.indexOf('notes');
  if (ch === -1 || item === -1) return kept;

  var cols = {}, cellNotes = {};
  INVENTORY_CATALOG_COLS.forEach(function(name) {
    cols[name] = hdr.indexOf(name);
    cellNotes[name] = cols[name] === -1 ? [] :
      sheet.getRange(1, cols[name] + 1, data.length, 1).getNotes();
  });

  for (var r = 1; r < data.length; r++) {
    var c = String(data[r][ch]).trim(), it = String(data[r][item]).trim();
    if (!c || !it) continue;

    var entry = { notes: notesCol === -1 ? '' : String(data[r][notesCol] || '').trim() };
    var any = !!entry.notes;
    INVENTORY_CATALOG_COLS.forEach(function(name) {
      var v = cols[name] === -1 ? '' : String(data[r][cols[name]] || '').trim();
      var fromCatalog = String((cellNotes[name][r] || [''])[0] || '').trim();
      entry[name] = (v && v !== fromCatalog) ? v : '';
      if (entry[name]) any = true;
    });
    if (any) kept[c.toLowerCase() + '|' + it.toLowerCase()] = entry;
  }
  return kept;
}

function buildInventorySheet(ss, txnSheet) {
  var sheetName = 'Inventory';
  var sheet = ss.getSheetByName(sheetName);
  var kept = sheet ? readInventoryManual(sheet) : {};
  var catalog = loadItemCatalog(ss);
  if (sheet) {
    var existingFilter = sheet.getFilter();
    if (existingFilter) existingFilter.remove();
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  var data = txnSheet.getRange(2, 1, txnSheet.getLastRow() - 1, 8).getValues();

  var items = {};
  var order = [];
  for (var i = 0; i < data.length; i++) {
    var charName = String(data[i][TXN.character - 1]).trim();
    var itemName = String(data[i][TXN.itemName - 1]).trim();
    var qty = Number(data[i][TXN.quantity - 1]) || 0;
    var cost = data[i][TXN.cost - 1];
    var weight = data[i][TXN.weight - 1];
    var source = String(data[i][TXN.source - 1]).trim();

    if (!charName || !itemName) continue;

    var key = charName.toLowerCase() + '|' + itemName.toLowerCase();
    if (!(key in items)) {
      items[key] = {
        character: charName, itemName: itemName,
        quantity: 0, cost: cost, weight: weight,
        firstSource: source,
      };
      order.push(key);
    }
    items[key].quantity += qty;
    if (cost && cost !== 'N/A' && Number(cost) > 0) items[key].cost = cost;
    if (weight && weight !== '' && Number(weight) > 0) items[key].weight = weight;
  }

  var rows = [];
  var NOTES_COL = 8;
  rows.push(['character', 'item_name', 'source', 'cost', 'weight',
             'quantity', 'total_value', 'notes'].concat(INVENTORY_CATALOG_COLS));
  var LINK_COL = NOTES_COL + 1 + INVENTORY_CATALOG_COLS.indexOf('link');

  for (var i = 0; i < order.length; i++) {
    var item = items[order[i]];
    if (item.quantity <= 0 && item.itemName.toLowerCase() !== 'gold') continue;

    var costNum = Number(item.cost) || 0;
    var totalValue = costNum * item.quantity;
    var sourceLabel = item.firstSource === 'game' ? 'Reward' : 'Bought';
    if (item.itemName.toLowerCase() === 'gold') sourceLabel = '';

    var key = order[i];
    var manual = kept[key] || {};
    var ref = catalog[item.itemName.toLowerCase()] || {};

    // Catalog-backed columns: the manual value, else the catalog one.
    // The last element is a helper holding the cell notes (the catalog
    // value each cell was filled from, see readInventoryManual); it is
    // cut off before writing.
    var values = [], notes = [];
    INVENTORY_CATALOG_COLS.forEach(function(name) {
      var fromCatalog = String(ref[name] || '').trim();
      values.push(manual[name] || fromCatalog);
      notes.push(manual[name] ? '' : fromCatalog);
    });

    rows.push([item.character, item.itemName, sourceLabel,
               item.cost, item.weight, item.quantity, totalValue,
               manual.notes || ''].concat(values, [notes]));
  }

  // Sort by character name (column 1), then item name (column 2)
  var dataRows = rows.slice(1); // everything except header
  dataRows.sort(function(a, b) {
    var charCompare = String(a[0]).toLowerCase().localeCompare(String(b[0]).toLowerCase());
    if (charCompare !== 0) return charCompare;
    return String(a[1]).toLowerCase().localeCompare(String(b[1]).toLowerCase());
  });
  rows = [rows[0]].concat(dataRows);
  
  var cellNotes = rows.slice(1).map(function(r) { return r.pop(); });

  if (rows.length <= 1) {
    sheet.getRange(1, 1, 1, rows[0].length).setValues([rows[0]]);
  } else {
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    sheet.getRange(2, NOTES_COL + 1, cellNotes.length, INVENTORY_CATALOG_COLS.length).setNotes(cellNotes);
  }

  // Format header
  var headerRange = sheet.getRange(1, 1, 1, rows[0].length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1B1B2F');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  // Highlight notes column
  if (rows.length > 1) {
    sheet.getRange(2, NOTES_COL, rows.length - 1, 1 + INVENTORY_CATALOG_COLS.length).setBackground('#FFF9E6');
  }

  // Alternate row colors
  for (var r = 2; r <= rows.length; r++) {
    var bg = (r % 2 === 0) ? '#F3E5F5' : '#FFFFFF';
    sheet.getRange(r, 1, 1, 7).setBackground(bg);
  }

  // Column widths
  for (var c = 1; c <= rows[0].length; c++) {
    sheet.autoResizeColumn(c);
    var width = sheet.getColumnWidth(c);
    sheet.setColumnWidth(c, width + 30);
  }
  sheet.setColumnWidth(8, 250);
  sheet.setColumnWidth(LINK_COL, 300);

  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, rows[0].length).createFilter();
  }

  sheet.setFrozenRows(1);
  sheet.setTabColor('#FF6600');
}


// ============================================================
// HELPERS
// ============================================================

function isApproved(val) {
  return val === true || val === 'TRUE' || val === 'True';
}