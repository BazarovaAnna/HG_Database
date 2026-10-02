// ============================================================
// PF1e Character Database — Feat Normalizer
// ============================================================
// Reads comma-separated feats and skills from the main sheet,
// parses them into normalized tables on separate sheets.
//
// CONVENTIONS EXPECTED:
//   Feats field:  "Power Attack, Cleave, Weapon Focus (longsword) [bonus], Toughness"
//                  [bonus] tag is optional, marks class bonus feats
// ============================================================


// ============================================================
// FEATS PARSER
// ============================================================

function parseFeats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var main = ss.getSheetByName(CONFIG.mainSheet);
  if (!main) throw new Error('Sheet "' + CONFIG.mainSheet + '" not found');

  var data = main.getDataRange().getValues();
  var headers = data[CONFIG.headerRow - 1];

  var col = findColumns(headers, {
    char:   CONFIG.charNameKey,
    feats:  CONFIG.featsKey,
    level:  CONFIG.levelKey,
    class1: CONFIG.classKeys[0],
    class2: CONFIG.classKeys[1],
    class3: CONFIG.classKeys[2],
    classFeatures:     CONFIG.classFeatures,
    specialAbilities:  CONFIG.specialAbilities,
  });

  // --- Save existing GM notes before clearing ---
  var gmNotes = {};
  var featSheet = ss.getSheetByName(CONFIG.featsSheet);
  if (featSheet && featSheet.getLastRow() > 1) {
    var existingData = featSheet.getDataRange().getValues();
    var existingHeaders = existingData[0];
    var charCol = -1, nameCol = -1, typeCol = -1, notesCol = -1;
    for (var c = 0; c < existingHeaders.length; c++) {
      var h = String(existingHeaders[c]).trim().toLowerCase();
      if (h === 'character') charCol = c;
      if (h === 'name') nameCol = c;
      if (h === 'type') typeCol = c;
      if (h === 'gm_notes') notesCol = c;
    }
    if (charCol !== -1 && nameCol !== -1 && typeCol !== -1 && notesCol !== -1) {
      for (var r = 1; r < existingData.length; r++) {
        var val = existingData[r][notesCol];
        if (val !== '' && val !== undefined && val !== null) {
          var key = String(existingData[r][charCol]).trim().toLowerCase() + '|' +
                    String(existingData[r][typeCol]).trim().toLowerCase() + '|' +
                    String(existingData[r][nameCol]).trim().toLowerCase();
          gmNotes[key] = val;
        }
      }
    }
  }

  var rows = [];
  rows.push([
    'character', 'level', 'classes',
    'type', 'name', 'source', 'gm_notes'
  ]);

  for (var i = CONFIG.dataStartRow - 1; i < data.length; i++) {
    var charName = data[i][col.char];
    var level = data[i][col.level];
    if (!charName || charName === '') continue;

    var classes = [data[i][col.class1], data[i][col.class2], data[i][col.class3]]
      .filter(function(c) { return c && c !== ''; })
      .join(' / ');

    // --- Parse feats ---
    var featsRaw = String(data[i][col.feats] || '');
    if (featsRaw.trim() !== '') {
      parseEntries(rows, gmNotes, charName, level, classes, featsRaw, 'feat');
    }

    // --- Parse class features ---
    var cfRaw = String(data[i][col.classFeatures] || '');
    if (cfRaw.trim() !== '') {
      parseEntries(rows, gmNotes, charName, level, classes, cfRaw, 'class feature');
    }

    // --- Parse special abilities ---
    var saRaw = String(data[i][col.specialAbilities] || '');
    if (saRaw.trim() !== '') {
      parseEntries(rows, gmNotes, charName, level, classes, saRaw, 'special ability');
    }
  }

  writeFeatSheet(ss, CONFIG.featsSheet, rows, '4A0E4E');
}


function parseEntries(rows, gmNotes, charName, level, classes, rawStr, type) {
  var entries = rawStr.split(',');
  for (var j = 0; j < entries.length; j++) {
    var raw = entries[j].trim();
    if (raw === '') continue;

    var source = '';
    var tagMatch = raw.match(/\[([^\]]+)\]/);
    if (tagMatch) {
      source = tagMatch[1].toLowerCase().trim();
      raw = raw.replace(/\s*\[[^\]]+\]/, '').trim();
    }

    var noteKey = String(charName).trim().toLowerCase() + '|' +
                  type + '|' + raw.toLowerCase();
    var note = (noteKey in gmNotes) ? gmNotes[noteKey] : '';

    rows.push([charName, level, classes, type, raw, source, note]);
  }
}

// ============================================================
// HELPERS
// ============================================================

function findColumns(headers, keyMap) {
  var result = {};
  for (var k in keyMap) {
    var idx = headers.indexOf(keyMap[k]);
    if (idx === -1) throw new Error('Column "' + keyMap[k] + '" not found in header row');
    result[k] = idx;
  }
  return result;
}

function writeFeatSheet(ss, sheetName, rows, tabColor) {
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    var existingFilter = sheet.getFilter();
    if (existingFilter) existingFilter.remove();
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  if (rows.length === 0) return;

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  // Format header
  var headerRange = sheet.getRange(1, 1, 1, rows[0].length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1B1B2F');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  // Highlight gm_notes column as editable
  if (rows.length > 1) {
    sheet.getRange(2, 7, rows.length - 1, 1).setBackground('#FFF9E6');
  }

  // Alternate row colors (skip gm_notes)
  for (var r = 2; r <= rows.length; r++) {
    var bg = (r % 2 === 0) ? '#F3E5F5' : '#FFFFFF';
    sheet.getRange(r, 1, 1, 6).setBackground(bg);
  }

  // Column widths
  for (var c = 1; c <= rows[0].length; c++) {
    sheet.autoResizeColumn(c);
    var width = sheet.getColumnWidth(c);
    sheet.setColumnWidth(c, width + 30);
  }
  sheet.setColumnWidth(7, 250);

  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, rows[0].length).createFilter();
  }

  sheet.setTabColor('#' + tabColor);
  sheet.setFrozenRows(1);
}
