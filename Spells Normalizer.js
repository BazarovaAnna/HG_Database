// ============================================================
// PF1e Character Database — Spell Normalizer
// ============================================================
// Parses cantrip and spell lists from the Database sheet into
// a normalized "Spells" sheet. One row per spell per character.
//
// Expected Database columns (row 4 keys):
//   cant_num, cant, spell1_num, spell1, ... spell9_num, spell9
//
// _num format: "known/prepared" e.g. "2/1" or "-/3" or "0/0"
// spell list format: comma-separated names
//
// Add to your existing Apps Script project.
// Update onOpen() and parseAll() to include parseSpells().
// ============================================================


function parseSpells() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var main = ss.getSheetByName('Database');
  if (!main) { SpreadsheetApp.getUi().alert('Sheet "Database" not found!'); return; }

  var data = main.getDataRange().getValues();
  var headers = data[3];

  var col = {};
  for (var c = 0; c < headers.length; c++) {
    var key = String(headers[c]).trim();
    if (key) col[key] = c;
  }

  function v(row, key) {
    if (!(key in col)) return '';
    var val = row[col[key]];
    return (val === undefined || val === null) ? '' : val;
  }

  function parseNum(numStr) {
  var s = String(numStr).trim();
  if (!s || s === '0;0' || s === '0/0') return { known: 0, prepared: 0 };
  // Support both ; and / separators
  var parts = s.split(/[;\/]/);
  var known = (parts[0].trim() === '-') ? -1 : (parseInt(parts[0]) || 0);
  var prepared = (parts[1] && parts[1].trim() === '-') ? -1 : (parseInt(parts[1]) || 0);
  return { known: known, prepared: prepared };
}

  // --- Save existing GM notes before clearing ---
  var gmNotes = {};
  var spellSheet = ss.getSheetByName('Spells');
  if (spellSheet && spellSheet.getLastRow() > 1) {
    var existingData = spellSheet.getDataRange().getValues();
    var existingHeaders = existingData[0];
    var charCol = -1, lvlCol = -1, nameCol = -1, notesCol = -1;
    for (var c = 0; c < existingHeaders.length; c++) {
      var h = String(existingHeaders[c]).trim().toLowerCase();
      if (h === 'character') charCol = c;
      if (h === 'spell_level') lvlCol = c;
      if (h === 'spell_name') nameCol = c;
      if (h === 'gm_notes') notesCol = c;
    }
    if (charCol !== -1 && lvlCol !== -1 && nameCol !== -1 && notesCol !== -1) {
      for (var r = 1; r < existingData.length; r++) {
        var val = existingData[r][notesCol];
        if (val !== '' && val !== undefined && val !== null) {
          var key = String(existingData[r][charCol]).trim().toLowerCase() + '|' +
                    String(existingData[r][lvlCol]) + '|' +
                    String(existingData[r][nameCol]).trim().toLowerCase();
          gmNotes[key] = val;
        }
      }
    }
  }

  var rows = [];
  rows.push([
    'character', 'level', 'cast_class',
    'spell_level', 'spell_name', 'known', 'prepared', 'gm_notes'
  ]);

  for (var i = 4; i < data.length; i++) {
    var row = data[i];
    var charName = String(v(row, 'char')).trim();
    if (!charName) continue;

    var level    = v(row, 'level');
    var castCls  = String(v(row, 'cast_class')).trim();
    var isCaster = String(v(row, 'is_caster')).trim().toLowerCase();

    if (!isCaster || isCaster === 'no' || isCaster === 'нет') continue;

    // Cantrips
    var cantNums = parseNum(v(row, 'cant_num'));
    var cantStr = String(v(row, 'cant')).trim();
    if (cantStr) {
      var cantrips = cantStr.split(',');
      for (var j = 0; j < cantrips.length; j++) {
        var name = cantrips[j].trim();
        if (name) {
          var noteKey = charName.toLowerCase() + '|0|' + name.toLowerCase();
          var note = (noteKey in gmNotes) ? gmNotes[noteKey] : '';
          rows.push([charName, level, castCls,
            0, name,
            cantNums.known === -1 ? '∞' : cantNums.known,
            cantNums.prepared === -1 ? '∞' : cantNums.prepared,
            note]);
        }
      }
    }

    // Spell levels 1-9
    for (var sl = 1; sl <= 9; sl++) {
      var numKey = 'spell' + sl + '_num';
      var listKey = 'spell' + sl;

      var nums = parseNum(v(row, numKey));
      var spellStr = String(v(row, listKey)).trim();
      if (!spellStr) continue;

      var spells = spellStr.split(',');
      for (var j = 0; j < spells.length; j++) {
        var name = spells[j].trim();
        if (name) {
          var noteKey = charName.toLowerCase() + '|' + sl + '|' + name.toLowerCase();
          var note = (noteKey in gmNotes) ? gmNotes[noteKey] : '';
          rows.push([charName, level, castCls,
            sl, name,
            nums.known === -1 ? '∞' : nums.known,
            nums.prepared === -1 ? '∞' : nums.prepared,
            note]);
        }
      }
    }
  }

  writeSpellSheet(ss, rows);
}

// ============================================================
// HELPERS
// ============================================================

function writeSpellSheet(ss, rows) {
  var sheetName = 'Spells';
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    var existingFilter = sheet.getFilter();
    if (existingFilter) existingFilter.remove();
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  if (rows.length <= 1) {
    sheet.getRange(1, 1, 1, rows[0].length).setValues([rows[0]]);
    var headerRange = sheet.getRange(1, 1, 1, rows[0].length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1B1B2F');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setTabColor('#4A0E4E');
    return;
  }

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  // Format header
  var headerRange = sheet.getRange(1, 1, 1, rows[0].length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1B1B2F');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  // Highlight gm_notes column as editable
  if (rows.length > 1) {
    sheet.getRange(2, 8, rows.length - 1, 1).setBackground('#FFF9E6');
  }

  // Alternate row colors (skip gm_notes column)
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

  // Align known/prepared to the right
  if (rows.length > 1) {
    sheet.getRange(2, 6, rows.length - 1, 2).setHorizontalAlignment('right');
  }

  // Filter
  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, rows[0].length).createFilter();
  }

  sheet.setFrozenRows(1);
  sheet.setTabColor('#4A0E4E');
}
