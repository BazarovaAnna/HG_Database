// ============================================================
// PF1e Character Database — Level-Up Processor v4
// ============================================================
//
// Reads column mappings from "Column matchups" sheet — no hardcoded
// column numbers. Uses loadColumnMappings() and getField() from
// the inventory processor script.
//
// Two-phase processing:
//   Phase 1 (UNDO): Revert processed level-ups that lost approval
//   Phase 2 (APPLY): Process newly approved level-ups
//
// REQUIRES:
//   - loadColumnMappings() and getField() from inventory processor
//   - isApproved() helper
//   - "Column matchups" sheet with Registration and LvlUP sections
//   - "Registration" sheet
//   - "LvlUP" sheet (form responses)
//   - "LevelUp Log" sheet (auto-created)
// ============================================================


// ============================================================
// LOG COLUMNS (we control these — always the same)
// ============================================================
var LOG = {
  timestamp:   1,
  processedAt: 2,
  character:   3,
  newLevel:    4,
  field:       5,
  regCol:      6,
  oldValue:    7,
  newValue:    8,
  gmNotes:     9,
};


// ============================================================
// MAIN ENTRY POINT
// ============================================================

function processLevelUps() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lvlSheet = ss.getSheetByName('LvlUP');
  var regSheet = ss.getSheetByName('Registration');
  if (!lvlSheet) { SpreadsheetApp.getUi().alert('Sheet "LvlUP" not found!'); return; }
  if (!regSheet) { SpreadsheetApp.getUi().alert('Sheet "Registration" not found!'); return; }

  var cols = loadColumnMappings(ss);
  var logSheet = getOrCreateLogSheet(ss);

  var undoReport = phaseUndo(ss, lvlSheet, regSheet, logSheet, cols);
  var applyReport = phaseApply(ss, lvlSheet, regSheet, logSheet, cols);

  // --- Summary ---
  var msg = 'Level-Up Processing Complete!\n\n';

  if (undoReport.reverted > 0 || undoReport.errors.length > 0) {
    msg += '--- UNDO PHASE ---\n';
    msg += '↩️ Reverted: ' + undoReport.reverted + ' level-ups\n';
    msg += '⚠️ Skipped fields: ' + undoReport.skippedFields + '\n';
    if (undoReport.details.length > 0) {
      for (var i = 0; i < undoReport.details.length; i++) {
        msg += '  ' + undoReport.details[i] + '\n';
      }
    }
    if (undoReport.errors.length > 0) {
      msg += '❌ Errors:\n';
      for (var i = 0; i < undoReport.errors.length; i++) {
        msg += '  • ' + undoReport.errors[i] + '\n';
      }
    }
    msg += '\n';
  }

  msg += '--- APPLY PHASE ---\n';
  msg += '✅ Processed: ' + applyReport.processed + '\n';
  msg += '⏭ Skipped (already done): ' + applyReport.skipped + '\n';
  if (applyReport.errors.length > 0) {
    msg += '❌ Errors:\n';
    for (var i = 0; i < applyReport.errors.length; i++) {
      msg += '  • ' + applyReport.errors[i] + '\n';
    }
  }

  msg += '\nSee "LevelUp Log" sheet for details.';
  SpreadsheetApp.getUi().alert(msg);
}


// ============================================================
// PHASE 1: UNDO
// ============================================================

function phaseUndo(ss, lvlSheet, regSheet, logSheet, cols) {
  var report = { reverted: 0, skippedFields: 0, errors: [], details: [] };
  var L = cols.lvlup;

  var lvlData = lvlSheet.getDataRange().getValues();
  var logData = logSheet.getDataRange().getValues();

  // Build set of processed timestamps from log
  var processedSet = {};
  for (var i = 1; i < logData.length; i++) {
    var key = String(logData[i][LOG.timestamp - 1]) + '|' +
              String(logData[i][LOG.character - 1]).trim().toLowerCase();
    processedSet[key] = true;
  }

  // Find processed rows that lost approval, grouped by character
  var unapproved = {};
  for (var i = 1; i < lvlData.length; i++) {
    var row = lvlData[i];
    var timestamp = String(getField(row, L, 'timestamp'));
    var charName = String(getField(row, L, 'charName') || '').trim();
    var approved = getField(row, L, 'approved');
    var newLevel = Number(getField(row, L, 'newLevel')) || 0;
    if (!charName) continue;

    var processKey = timestamp + '|' + charName.toLowerCase();
    var isProcessed = processedSet[processKey] || false;

    if (isProcessed && !isApproved(approved)) {
      if (!unapproved[charName.toLowerCase()]) unapproved[charName.toLowerCase()] = [];
      unapproved[charName.toLowerCase()].push({
        lvlSheetRow: i + 1,
        timestamp: timestamp,
        charName: charName,
        newLevel: newLevel,
      });
    }
  }

  // For each character with unapproved levels, cascade
  for (var charKey in unapproved) {
    var entries = unapproved[charKey];
    var charName = entries[0].charName;

    // Find earliest unapproved level
    var earliestLevel = 999;
    for (var i = 0; i < entries.length; i++) {
      if (entries[i].newLevel < earliestLevel) earliestLevel = entries[i].newLevel;
    }

    // Collect ALL processed level-ups at or above earliest level
    var toRevert = [];
    for (var i = 1; i < lvlData.length; i++) {
      var row = lvlData[i];
      var ts = String(getField(row, L, 'timestamp'));
      var cn = String(getField(row, L, 'charName') || '').trim();
      var nl = Number(getField(row, L, 'newLevel')) || 0;
      if (cn.toLowerCase() !== charKey) continue;

      var pk = ts + '|' + cn.toLowerCase();
      if (!processedSet[pk]) continue;
      if (nl >= earliestLevel) {
        toRevert.push({ timestamp: ts, newLevel: nl, lvlSheetRow: i + 1 });
      }
    }

    // Sort descending — revert highest first
    toRevert.sort(function(a, b) { return b.newLevel - a.newLevel; });

    // Revert each one
    var revertedLevels = [];
    var cascadedLevels = [];
    for (var i = 0; i < toRevert.length; i++) {
      var entry = toRevert[i];
      try {
        var result = revertLevelUp(regSheet, logSheet, logData, entry.timestamp, charName, cols);
        report.reverted++;
        report.skippedFields += result.skipped;
        revertedLevels.push(entry.newLevel);

        // Uncheck approve — track if it was still checked (cascade)
        var wasApproved = lvlSheet.getRange(entry.lvlSheetRow, L.approved).getValue();
        lvlSheet.getRange(entry.lvlSheetRow, L.approved).setValue(false);
        if (isApproved(wasApproved)) {
          cascadedLevels.push(entry.newLevel);
        }

      } catch (e) {
        report.errors.push(charName + ' (lvl ' + entry.newLevel + '): ' + e.message);
      }
    }

    // Build detail message
    if (revertedLevels.length > 0) {
      var detail = charName + ': reverted levels ' + revertedLevels.join(', ') +
        ' (triggered by unchecked level ' + earliestLevel + ')';
      if (cascadedLevels.length > 0) {
        detail += '\n    ⚠️ Auto-unchecked approved levels: ' + cascadedLevels.join(', ');
      }
      report.details.push(detail);
    }
  }

  return report;
}


function revertLevelUp(regSheet, logSheet, logData, timestamp, charName, cols) {
  var result = { skipped: 0 };
  var R = cols.registration;

  // Find all log rows for this timestamp + character
  var logRows = [];
  for (var i = 1; i < logData.length; i++) {
    var logTs = String(logData[i][LOG.timestamp - 1]);
    var logChar = String(logData[i][LOG.character - 1]).trim();
    if (logTs === timestamp && logChar.toLowerCase() === charName.toLowerCase()) {
      logRows.push({
        sheetRow: i + 1,
        field: String(logData[i][LOG.field - 1]).trim(),
        regCol: Number(logData[i][LOG.regCol - 1]) || 0,
        oldValue: logData[i][LOG.oldValue - 1],
        newValue: logData[i][LOG.newValue - 1],
      });
    }
  }

  if (logRows.length === 0) {
    throw new Error('No log entries found for timestamp ' + timestamp);
  }

  // Find character row in Registration
  var regData = regSheet.getDataRange().getValues();
  var regRow = findCharacterRow(regData, charName, R);
  if (regRow === -1) throw new Error('Character not found in Registration');

  // Revert each change
  for (var i = 0; i < logRows.length; i++) {
    var entry = logRows[i];
    if (entry.field === '(no changes)') continue;
    if (entry.regCol <= 0) { result.skipped++; continue; }

    var currentValue = String(regSheet.getRange(regRow, entry.regCol).getValue());
    var expectedValue = String(entry.newValue);

    var match = false;
    var currentNum = Number(currentValue);
    var expectedNum = Number(expectedValue);
    if (!isNaN(currentNum) && !isNaN(expectedNum) && currentNum === expectedNum) {
      match = true;
    } else if (currentValue === expectedValue) {
      match = true;
    }

    if (match) {
      regSheet.getRange(regRow, entry.regCol).setValue(entry.oldValue);
    } else {
      result.skipped++;
    }
  }

  // Delete log entries (reverse order)
  var rowsToDelete = logRows.map(function(r) { return r.sheetRow; }).sort(function(a, b) { return b - a; });
  for (var i = 0; i < rowsToDelete.length; i++) {
    logSheet.deleteRow(rowsToDelete[i]);
  }

  return result;
}


// ============================================================
// PHASE 2: APPLY
// ============================================================

function phaseApply(ss, lvlSheet, regSheet, logSheet, cols) {
  var report = { processed: 0, skipped: 0, errors: [] };
  var L = cols.lvlup;
  var R = cols.registration;

  var lvlData = lvlSheet.getDataRange().getValues();

  // Reload log after undo phase
  var processedSet = getProcessedSet(logSheet);

  for (var i = 1; i < lvlData.length; i++) {
    var row = lvlData[i];
    var timestamp = String(getField(row, L, 'timestamp'));
    var charName = String(getField(row, L, 'charName') || '').trim();
    var approved = getField(row, L, 'approved');

    if (!isApproved(approved)) continue;
    if (!charName) continue;

    var processKey = timestamp + '|' + charName.toLowerCase();
    if (processedSet[processKey]) { report.skipped++; continue; }

    var regData = regSheet.getDataRange().getValues();
    var regRow = findCharacterRow(regData, charName, R);
    if (regRow === -1) {
      report.errors.push(charName + ': not found in Registration');
      continue;
    }

    try {
      var changes = applyLevelUp(regSheet, row, regRow, L, R);
      var gmNotes = String(getField(row, L, 'gmNotes') || '').trim();
      var newLevel = getField(row, L, 'newLevel');
      logLevelUp(logSheet, timestamp, charName, newLevel, changes, gmNotes);
      report.processed++;
    } catch (e) {
      report.errors.push(charName + ': ' + e.message);
    }
  }

  return report;
}


function applyLevelUp(regSheet, lvlRow, regRow, L, R) {
  var changes = [];

  // --- 1. Replace input_lvl ---
  var newLevel = getField(lvlRow, L, 'newLevel');
  if (newLevel && newLevel !== '') {
    changes.push(setCell(regSheet, regRow, R.inputLvl, newLevel, 'input_lvl'));
  }

  // --- 2. Increment class level ---
  var classToLevel = String(getField(lvlRow, L, 'classToLevel') || '').trim();
  if (classToLevel) {
    var classPairs = [
      { nameKey: 'class1', lvlKey: 'class1Lvl', archKey: 'class1Arch' },
      { nameKey: 'class2', lvlKey: 'class2Lvl', archKey: 'class2Arch' },
      { nameKey: 'class3', lvlKey: 'class3Lvl', archKey: 'class3Arch' },
    ];
    var classFound = false;

    for (var cp = 0; cp < classPairs.length; cp++) {
      var nameCol = R[classPairs[cp].nameKey];
      var lvlCol = R[classPairs[cp].lvlKey];
      if (!nameCol || !lvlCol) continue;

      var existing = String(regSheet.getRange(regRow, nameCol).getValue() || '').trim();
      if (existing.toLowerCase() === classToLevel.toLowerCase()) {
        var oldLvl = Number(regSheet.getRange(regRow, lvlCol).getValue()) || 0;
        changes.push(setCell(regSheet, regRow, lvlCol, oldLvl + 1, classPairs[cp].lvlKey));
        classFound = true;
        break;
      }
    }

    // New class — find first empty slot
    if (!classFound) {
      for (var cp = 0; cp < classPairs.length; cp++) {
        var nameCol = R[classPairs[cp].nameKey];
        var lvlCol = R[classPairs[cp].lvlKey];
        if (!nameCol || !lvlCol) continue;

        var existing = String(regSheet.getRange(regRow, nameCol).getValue() || '').trim();
        if (!existing) {
          changes.push(setCell(regSheet, regRow, nameCol, classToLevel, classPairs[cp].nameKey));
          changes.push(setCell(regSheet, regRow, lvlCol, 1, classPairs[cp].lvlKey));

          // Also set archetype if provided
          var newArch = String(getField(lvlRow, L, 'newArch') || '').trim();
          var archCol = R[classPairs[cp].archKey];
          if (newArch && archCol) {
            changes.push(setCell(regSheet, regRow, archCol, newArch, classPairs[cp].archKey));
          }
          classFound = true;
          break;
        }
      }
    }
  }

  // --- 3. HP ---
  var hpRolled = Number(getField(lvlRow, L, 'hpRolled')) || 0;
  if (hpRolled > 0) {
    var oldHP = Number(regSheet.getRange(regRow, R.hpTotal).getValue()) || 0;
    changes.push(setCell(regSheet, regRow, R.hpTotal, oldHP + hpRolled, 'hpTotal'));
  }

  // --- 4. Favored class bonus ---
  var fcBonus = String(getField(lvlRow, L, 'fcBonus') || '').trim();
  if (fcBonus) {
    var fcMap = {
      'hp':     { key: 'fcHP' },
      'skill':  { key: 'fcSkill' },
      'racial': { key: 'fcRacial' },
    };
    var fc = fcMap[fcBonus.toLowerCase()];
    if (fc && R[fc.key]) {
      var oldFC = Number(regSheet.getRange(regRow, R[fc.key]).getValue()) || 0;
      changes.push(setCell(regSheet, regRow, R[fc.key], oldFC + 1, fc.key));
    }
  }

  // --- 5. Ability score increase ---
  // Only the choice is recorded. The raw score stays the base value:
  // Database adds the increases itself by counting abilityInc columns.
  var abilityInc = String(getField(lvlRow, L, 'abilityInc') || '').trim();
  if (abilityInc && abilityInc.toLowerCase() !== 'none') {
    var lvlNum = Number(newLevel) || 0;

    // Map milestone levels to Registration column keys
    var incKeyMap = { 4: 'abilityInc4', 8: 'abilityInc8', 12: 'abilityInc12',
                      16: 'abilityInc16', 20: 'abilityInc20' };
    var incKey = incKeyMap[lvlNum];
    if (incKey && R[incKey]) {
      changes.push(setCell(regSheet, regRow, R[incKey], abilityInc, incKey));
    } else {
      // regCol 0: nothing to revert on undo
      changes.push({ field: 'abilityInc', regCol: 0, old: '',
                     new: 'IGNORED — ' + abilityInc + ' at level ' + lvlNum + ' (not 4/8/12/16/20)' });
    }
  }

  // --- 6. New favored class ---
  var newFavClass = String(getField(lvlRow, L, 'newFavClass') || '').trim();
  if (newFavClass) {
    changes.push(setCell(regSheet, regRow, R.favoredClass, newFavClass, 'favoredClass'));
  }

  // --- 7. Append text fields ---
  var appendFields = [
    { lvlKey: 'newFeats',       regKey: 'feats' },
    { lvlKey: 'newClassFeat',   regKey: 'classFeatures' },
    { lvlKey: 'newSpecialAbil', regKey: 'specialAbil' },
    { lvlKey: 'newTraits',      regKey: 'traits' },
    { lvlKey: 'newDrawbacks',   regKey: 'drawbacks' },
  ];
  for (var af = 0; af < appendFields.length; af++) {
    var newText = String(getField(lvlRow, L, appendFields[af].lvlKey) || '').trim();
    var regCol = R[appendFields[af].regKey];
    if (newText && regCol) {
      changes.push(appendCell(regSheet, regRow, regCol, newText, appendFields[af].regKey));
    }
  }

  // --- 8. Merge skills ---
  var newSkillRanks = String(getField(lvlRow, L, 'newSkillRanks') || '').trim();
  if (newSkillRanks && R.skills) {
    changes.push(mergeSkills(regSheet, regRow, R.skills, newSkillRanks));
  }

  // --- 9. Update is_caster ---
  var isCaster = String(getField(lvlRow, L, 'isCaster') || '').trim();
  if (isCaster && R.isCaster) {
    changes.push(setCell(regSheet, regRow, R.isCaster, isCaster, 'isCaster'));
  }

  // --- 10. Spells: replace _num, append lists ---
  if (R.cantNum) {
    var newCantNum = String(getField(lvlRow, L, 'newCantNum') || '').trim();
    if (newCantNum) {
      changes.push(setCell(regSheet, regRow, R.cantNum, newCantNum, 'cantNum'));
    }
  }
  if (R.cant) {
    var newCant = String(getField(lvlRow, L, 'newCant') || '').trim();
    if (newCant) {
      changes.push(appendCell(regSheet, regRow, R.cant, newCant, 'cant'));
    }
  }

  // Spell levels 1-9: derive column positions from spell1Num/spell1
  if (R.spell1Num && R.spell1) {
    for (var sl = 1; sl <= 9; sl++) {
      var regNumCol = R.spell1Num + (sl - 1) * 2;
      var regListCol = R.spell1 + (sl - 1) * 2;

      // LvlUP spell columns: newSpell1Num at col 19, newSpell1 at col 20
      // Pattern: newSpellNNum = newSpell1Num + (N-1)*2
      var lvlNumKey = 'newSpell' + sl + 'Num';
      var lvlListKey = 'newSpell' + sl;

      // Try dynamic key first, fall back to offset calculation
      var lvlNumCol = L[lvlNumKey];
      var lvlListCol = L[lvlListKey];

      // If not in matchups, calculate from spell1 pattern
      if (!lvlNumCol && L.newSpell1Num) lvlNumCol = L.newSpell1Num + (sl - 1) * 2;
      if (!lvlListCol && L.newSpell1) lvlListCol = L.newSpell1 + (sl - 1) * 2;

      if (lvlNumCol) {
        var newNum = String(lvlRow[lvlNumCol - 1] || '').trim();
        if (newNum) {
          changes.push(setCell(regSheet, regRow, regNumCol, newNum, 'spell' + sl + 'Num'));
        }
      }
      if (lvlListCol) {
        var newList = String(lvlRow[lvlListCol - 1] || '').trim();
        if (newList) {
          changes.push(appendCell(regSheet, regRow, regListCol, newList, 'spell' + sl));
        }
      }
    }
  }

  // --- 11. Spell retraining ---
  var retrained = String(getField(lvlRow, L, 'spellsRetrained') || '').trim();
  if (retrained && R.cant) {
    var retrainChanges = processRetraining(regSheet, regRow, retrained, R);
    for (var rc = 0; rc < retrainChanges.length; rc++) {
      changes.push(retrainChanges[rc]);
    }
  }

  return changes;
}


// ============================================================
// CELL OPERATION HELPERS
// ============================================================

function setCell(sheet, row, col, newValue, fieldName) {
  var oldValue = sheet.getRange(row, col).getValue();
  sheet.getRange(row, col).setValue(newValue);
  return { field: fieldName, regCol: col, old: String(oldValue), new: String(newValue) };
}

function appendCell(sheet, row, col, newText, fieldName) {
  var oldValue = String(sheet.getRange(row, col).getValue() || '').trim();
  var merged = oldValue ? oldValue + ', ' + newText : newText;
  sheet.getRange(row, col).setValue(merged);
  return { field: fieldName, regCol: col, old: oldValue, new: merged };
}

function mergeSkills(sheet, row, col, newSkillsStr) {
  var oldStr = String(sheet.getRange(row, col).getValue() || '').trim();

  var skillMap = {};
  var order = [];
  if (oldStr) {
    var oldEntries = oldStr.split(',');
    for (var i = 0; i < oldEntries.length; i++) {
      var parsed = parseLvlSkillEntry(oldEntries[i]);
      if (parsed) {
        skillMap[parsed.name.toLowerCase()] = parsed.ranks;
        order.push(parsed.name.toLowerCase());
      }
    }
  }

  var newEntries = newSkillsStr.split(',');
  for (var i = 0; i < newEntries.length; i++) {
    var parsed = parseLvlSkillEntry(newEntries[i]);
    if (parsed) {
      var key = parsed.name.toLowerCase();
      if (key in skillMap) {
        skillMap[key] += parsed.ranks;
      } else {
        skillMap[key] = parsed.ranks;
        order.push(parsed.name.toLowerCase());
      }
    }
  }

  var parts = [];
  for (var i = 0; i < order.length; i++) {
    parts.push(order[i] + ' ' + skillMap[order[i].toLowerCase()]);
  }
  var merged = parts.join(', ');

  sheet.getRange(row, col).setValue(merged);
  return { field: 'skills', regCol: col, old: oldStr, new: merged };
}

function parseLvlSkillEntry(entry) {
  var s = entry.trim();
  if (!s) return null;
  var parts = s.split(/\s+/);
  if (parts.length === 0) return null;

  var lastToken = parts[parts.length - 1];
  var ranks = 1; // default: no number means +1
  var nameParts = parts;

  if (!isNaN(lastToken) && lastToken !== '' && parts.length > 1) {
    ranks = parseInt(lastToken, 10) || 1;
    nameParts = parts.slice(0, -1);
  }

  var name = nameParts.join(' ').trim();
  if (!name) return null;

  return { name: name, ranks: ranks };
}

function processRetraining(sheet, row, retrainStr, R) {
  var changes = [];
  var swaps = retrainStr.split(',');

  // Build list of all spell list columns
  var spellCols = [];
  if (R.cant) spellCols.push(R.cant);
  if (R.spell1) {
    for (var sl = 1; sl <= 9; sl++) {
      spellCols.push(R.spell1 + (sl - 1) * 2);
    }
  }

  for (var i = 0; i < swaps.length; i++) {
    var swap = swaps[i].trim();
    // Support both "old->new" and "-old\nnew" format
    var parts = swap.split('->');
    if (parts.length !== 2) {
      // Try newline format: "-OldSpell\nNewSpell"
      parts = swap.split('\n');
      if (parts.length === 2) {
        parts[0] = parts[0].replace(/^-/, '').trim();
        parts[1] = parts[1].trim();
      } else {
        continue;
      }
    }

    var oldSpell = parts[0].trim();
    var newSpell = parts[1].trim();
    if (!oldSpell || !newSpell) continue;

    var found = false;
    for (var sc = 0; sc < spellCols.length; sc++) {
      var cellValue = String(sheet.getRange(row, spellCols[sc]).getValue() || '').trim();
      if (!cellValue) continue;

      var spellList = cellValue.split(',').map(function(s) { return s.trim(); });
      var idx = -1;
      for (var j = 0; j < spellList.length; j++) {
        if (spellList[j].toLowerCase() === oldSpell.toLowerCase()) {
          idx = j;
          break;
        }
      }

      if (idx !== -1) {
        var oldList = cellValue;
        spellList[idx] = newSpell;
        var newList = spellList.join(', ');
        sheet.getRange(row, spellCols[sc]).setValue(newList);
        changes.push({ field: 'spell_retrain', regCol: spellCols[sc], old: oldList, new: newList });
        found = true;
        break;
      }
    }

    if (!found) {
      changes.push({ field: 'spell_retrain', regCol: 0, old: oldSpell, new: 'NOT FOUND — ' + newSpell });
    }
  }

  return changes;
}


// ============================================================
// LOOKUP HELPERS
// ============================================================

function findCharacterRow(regData, charName, R) {
  var target = charName.toLowerCase().trim();
  for (var i = 1; i < regData.length; i++) {
    var name = String(regData[i][R.charName - 1] || '').toLowerCase().trim();
    if (name === target) return i + 1; // 1-indexed for sheet
  }
  return -1;
}

function getProcessedSet(logSheet) {
  var set = {};
  if (logSheet.getLastRow() <= 1) return set;
  var data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < data.length; i++) {
    var key = String(data[i][0]) + '|' + String(data[i][2]).trim().toLowerCase();
    set[key] = true;
  }
  return set;
}


// ============================================================
// LOGGING
// ============================================================

function getOrCreateLogSheet(ss) {
  var sheet = ss.getSheetByName('LevelUp Log');
  if (!sheet) {
    sheet = ss.insertSheet('LevelUp Log');
    var headers = ['Timestamp', 'Processed At', 'Character', 'New Level',
                   'Field', 'Reg Col', 'Old Value', 'New Value', 'GM Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1B1B2F');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setTabColor('#006400');
  }
  return sheet;
}

function logLevelUp(logSheet, timestamp, charName, newLevel, changes, gmNotes) {
  var processedAt = new Date().toISOString();
  var rows = [];

  if (changes.length === 0) {
    rows.push([timestamp, processedAt, charName, newLevel, '(no changes)', 0, '', '', gmNotes]);
  } else {
    for (var i = 0; i < changes.length; i++) {
      rows.push([
        timestamp, processedAt, charName, newLevel,
        changes[i].field,
        changes[i].regCol,
        changes[i].old,
        changes[i].new,
        (i === 0) ? gmNotes : '',
      ]);
    }
  }

  var startRow = logSheet.getLastRow() + 1;
  var range = logSheet.getRange(startRow, 1, rows.length, rows[0].length);
  range.setValues(rows);
}