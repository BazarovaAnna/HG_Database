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
      var txns = parseGameForm(row, G);
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
      var result = parseInventoryForm(row, I, inventory, charName);
      writeTxnRows(txnSheet, result.txns);
      for (var w = 0; w < result.warnings.length; w++) {
        warnings.push(charName + ': ' + result.warnings[w]);
      }
      for (var t = 0; t < result.txns.length; t++) {
        updateInventoryMap(inventory, result.txns[t]);
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

function parseGameForm(row, G) {
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
        var parsed = parseItemFull(items[it].trim());
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

function parseInventoryForm(row, I, inventory, charName) {
  var txns = [];
  var warnings = [];
  var timestamp = String(getField(row, I, 'timestamp'));
  var comment = String(getField(row, I, 'comment') || '').trim();

  // --- Items bought ---
  var boughtStr = String(getField(row, I, 'itemsBought') || '').trim();
  if (boughtStr) {
    var boughtItems = boughtStr.split(',');
    for (var i = 0; i < boughtItems.length; i++) {
      var parsed = parseItemFull(boughtItems[i].trim());
      if (!parsed) continue;

      var qty = Math.abs(parsed.quantity);
      txns.push({
        timestamp: timestamp, source: 'inventory', character: charName,
        itemName: parsed.name, quantity: qty,
        cost: parsed.cost, weight: parsed.weight,
        notes: comment,
      });

      var totalCost = parsed.cost * qty;
      if (totalCost > 0) {
        txns.push({
          timestamp: timestamp, source: 'inventory', character: charName,
          itemName: 'gold', quantity: -totalCost, cost: 1, weight: '',
          notes: 'Bought ' + (qty > 1 ? qty + ' ' : '') + parsed.name,
        });
      }
    }
  }

  // --- Items sold ---
  var soldStr = String(getField(row, I, 'itemsSold') || '').trim();
  if (soldStr) {
    var soldItems = soldStr.split(',');
    for (var i = 0; i < soldItems.length; i++) {
      var parsed = parseItemSimple(soldItems[i].trim());
      if (!parsed) continue;

      var lookup = lookupItem(inventory, charName, parsed.name);
      if (!lookup.found) {
        warnings.push('"' + parsed.name + '" not found in inventory — cost set to N/A');
      }

      var qty = Math.abs(parsed.quantity);
      txns.push({
        timestamp: timestamp, source: 'inventory', character: charName,
        itemName: parsed.name, quantity: -qty,
        cost: lookup.cost, weight: lookup.weight,
        notes: comment,
      });

      if (lookup.found && Number(lookup.cost) > 0) {
        var sellGold = Math.floor(Number(lookup.cost) * qty / 2);
        txns.push({
          timestamp: timestamp, source: 'inventory', character: charName,
          itemName: 'gold', quantity: sellGold, cost: 1, weight: '',
          notes: 'Sold ' + (qty > 1 ? qty + ' ' : '') + parsed.name,
        });
      }
    }
  }

  // --- Items gifted ---
  var giftedStr = String(getField(row, I, 'itemsGifted') || '').trim();
  if (giftedStr) {
    var giftedItems = giftedStr.split(',');
    for (var i = 0; i < giftedItems.length; i++) {
      var parsed = parseItemSimple(giftedItems[i].trim());
      if (!parsed) continue;

      var lookup = lookupItem(inventory, charName, parsed.name);
      if (!lookup.found) {
        warnings.push('"' + parsed.name + '" not found in inventory — cost set to N/A');
      }

      txns.push({
        timestamp: timestamp, source: 'inventory', character: charName,
        itemName: parsed.name, quantity: -Math.abs(parsed.quantity),
        cost: lookup.cost, weight: lookup.weight,
        notes: comment,
      });
    }
  }

  // --- Items got ---
  var gotStr = String(getField(row, I, 'itemsGot') || '').trim();
  if (gotStr) {
    var gotItems = gotStr.split(',');
    for (var i = 0; i < gotItems.length; i++) {
      var parsed = parseItemFull(gotItems[i].trim());
      if (!parsed) continue;

      txns.push({
        timestamp: timestamp, source: 'inventory', character: charName,
        itemName: parsed.name, quantity: Math.abs(parsed.quantity),
        cost: parsed.cost, weight: parsed.weight,
        notes: comment,
      });
    }
  }

  // --- Gold changes ---
  var goldChange = Number(getField(row, I, 'goldDelta')) || 0;
  if (goldChange !== 0) {
    txns.push({
      timestamp: timestamp, source: 'inventory', character: charName,
      itemName: 'gold', quantity: goldChange, cost: 1, weight: '',
      notes: comment,
    });
  }

  return { txns: txns, warnings: warnings };
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

  return { name: name, quantity: quantity, cost: cost, weight: weight };
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
    return { found: true, cost: map[key].cost || 0, weight: map[key].weight || '' };
  }
  return { found: false, cost: 'N/A', weight: '' };
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

function buildInventorySheet(ss, txnSheet) {
  var sheetName = 'Inventory';
  var sheet = ss.getSheetByName(sheetName);
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
  rows.push(['character', 'item_name', 'source', 'cost', 'weight',
             'quantity', 'total_value', 'notes']);

  for (var i = 0; i < order.length; i++) {
    var item = items[order[i]];
    if (item.quantity <= 0 && item.itemName.toLowerCase() !== 'gold') continue;

    var costNum = Number(item.cost) || 0;
    var totalValue = costNum * item.quantity;
    var sourceLabel = item.firstSource === 'game' ? 'Reward' : 'Bought';
    if (item.itemName.toLowerCase() === 'gold') sourceLabel = '';

    rows.push([item.character, item.itemName, sourceLabel,
               item.cost, item.weight, item.quantity, totalValue, '']);
  }

  // Sort by character name (column 1), then item name (column 2)
  var dataRows = rows.slice(1); // everything except header
  dataRows.sort(function(a, b) {
    var charCompare = String(a[0]).toLowerCase().localeCompare(String(b[0]).toLowerCase());
    if (charCompare !== 0) return charCompare;
    return String(a[1]).toLowerCase().localeCompare(String(b[1]).toLowerCase());
  });
  rows = [rows[0]].concat(dataRows);
  
  if (rows.length <= 1) {
    sheet.getRange(1, 1, 1, rows[0].length).setValues([rows[0]]);
  } else {
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  }

  // Format header
  var headerRange = sheet.getRange(1, 1, 1, rows[0].length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1B1B2F');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  // Highlight notes column
  if (rows.length > 1) {
    sheet.getRange(2, 8, rows.length - 1, 1).setBackground('#FFF9E6');
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


// ============================================================
// UPDATED onOpen
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PF1e Tools')
    .addItem('Refresh ALL normalized tables', 'parseAll')
    .addSeparator()
    .addItem('Refresh Feats only', 'parseFeats')
    .addItem('Refresh Skills only', 'parseSkills')
    .addItem('Refresh Spells only', 'parseSpells')
    .addSeparator()
    .addItem('Process Level-Ups', 'processLevelUps')
    .addSeparator()
    .addItem('Process Game Sessions', 'processGameSessions')
    .addItem('Process Inventory', 'processInventory')
    .addItem('Rebuild Inventory Sheet', 'rebuildInventory')
    .addSeparator()
    .addItem('Find data Discrepancies', 'findDiscrepancies')
    .addToUi();
}