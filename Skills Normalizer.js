// ============================================================
// PF1e Character Database — Skill Normalizer
// ============================================================
// Reads comma-separated skills from the main sheet,
// parses them into normalized tables on separate sheets.
//
// CONVENTIONS EXPECTED:
//   Skills field:  "acrobatics 1, stealth 5, perception 7"
//                  format: "skill_name ranks" (last token is the number)
// ============================================================


// ============================================================
// SKILLS PARSER
// ============================================================

function parseSkills() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var main = ss.getSheetByName('Database');
  var ref = ss.getSheetByName('Reference tables');
  if (!main) { SpreadsheetApp.getUi().alert('Sheet "Database" not found!'); return; }
  if (!ref) { SpreadsheetApp.getUi().alert('Sheet "Reference tables" not found!'); return; }

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
  function vNum(row, key) {
    var val = v(row, key);
    if (val === '' || val === undefined) return 0;
    return Number(val) || 0;
  }

  var abilityModKeys = {
    'STR': 'str_mod', 'DEX': 'dex_mod', 'CON': 'con_mod',
    'INT': 'int_mod', 'WIS': 'wis_mod', 'CHA': 'cha_mod',
  };

  var skillRef = loadSkillTable(ref);

  // --- Save existing GM edits before clearing ---
  var gmEdits = {};
  var skillSheet = ss.getSheetByName('Skills');
  if (skillSheet && skillSheet.getLastRow() > 1) {
    var existingData = skillSheet.getDataRange().getValues();
    var existingHeaders = existingData[0];
    // Find column indices by header name
    var gmCol = -1, charCol = -1, nameCol = -1;
    for (var c = 0; c < existingHeaders.length; c++) {
      var h = String(existingHeaders[c]).trim().toLowerCase();
      if (h === 'character') charCol = c;
      if (h === 'skill_name') nameCol = c;
      if (h === 'misc_mod') gmCol = c;
    }
    if (gmCol !== -1 && charCol !== -1 && nameCol !== -1) {
      for (var r = 1; r < existingData.length; r++) {
        var val = existingData[r][gmCol];
        if (val !== '' && val !== 0 && val !== undefined && val !== null) {
          var key = String(existingData[r][charCol]).trim().toLowerCase() + '|' +
                    String(existingData[r][nameCol]).trim().toLowerCase();
          gmEdits[key] = val;
        }
      }
    }
  }

  // --- Build rows (without sum — that will be a formula) ---
  // Columns: character, class_level, skill_name, ability,
  //          is_class, ability_score, ranks, class_mod, gm_mod, sum
  var rows = [];
  rows.push([
    'character', 'class_level', 'skill_name', 'ability',
    'is_class', 'ability_score', 'ranks', 'class_mod', 'misc_mod', 'sum'
  ]);

  for (var i = 4; i < data.length; i++) {
    var row = data[i];
    var charName = String(v(row, 'char')).trim();
    if (!charName) continue;

    var classLevelStr = String(v(row, 'multiclass')).trim();

    var charClasses = [];
    var cls1 = String(v(row, 'class_1')).trim();
    var cls2 = String(v(row, 'class_2')).trim();
    var cls3 = String(v(row, 'class_3')).trim();
    if (cls1) charClasses.push(cls1);
    if (cls2) charClasses.push(cls2);
    if (cls3) charClasses.push(cls3);

    var abilityMods = {};
    for (var ab in abilityModKeys) {
      abilityMods[ab] = vNum(row, abilityModKeys[ab]);
    }

    var skillsRaw = String(v(row, 'skills') || '');
    if (skillsRaw.trim() === '') continue;

    var skills = skillsRaw.split(',');
    for (var j = 0; j < skills.length; j++) {
      var entry = skills[j].trim();
      if (entry === '') continue;

      var parts = entry.split(/\s+/);
      var ranks = 0;
      var skillName = entry;

      var lastToken = parts[parts.length - 1];
      if (!isNaN(lastToken) && lastToken !== '') {
        ranks = parseInt(lastToken, 10);
        skillName = parts.slice(0, -1).join(' ');
      }

      skillName = skillName.toLowerCase().trim();

      var refEntry = findSkillRef(skillName, skillRef);
      var ability = refEntry ? refEntry.ability : '???';
      var abilityScore = (ability in abilityMods) ? abilityMods[ability] : 0;

      var isClassSkill = 'No';
      if (refEntry && refEntry.classes.length > 0) {
        for (var ci = 0; ci < charClasses.length; ci++) {
          for (var si = 0; si < refEntry.classes.length; si++) {
            if (charClasses[ci].toLowerCase() === refEntry.classes[si].toLowerCase()) {
              isClassSkill = 'Yes';
              break;
            }
          }
          if (isClassSkill === 'Yes') break;
        }
      }

      var classMod = (isClassSkill === 'Yes' && ranks > 0) ? 3 : 0;

      // Restore GM edit if exists
      var editKey = charName.toLowerCase() + '|' + skillName;
      var gmMod = (editKey in gmEdits) ? gmEdits[editKey] : 0;

      // Sum is placeholder — will be replaced by formula
      rows.push([
        charName, classLevelStr, skillName, ability,
        isClassSkill, abilityScore, ranks, classMod, gmMod, 0
      ]);
    }
  }

  // --- Write sheet ---
  writeSkillSheet(ss, rows);
}

// ============================================================
// HELPERS
// ============================================================

function loadSkillTable(refSheet) {
  var data = refSheet.getDataRange().getValues();
  var skillTable = [];

  for (var i = 0; i < data.length; i++) {
    var rowStr = data[i].join('|');
    if (rowStr.indexOf('Skill') !== -1 && rowStr.indexOf('Ability') !== -1) {
      var hdr = data[i];
      var sCols = {};
      for (var c = 0; c < hdr.length; c++) {
        var h = String(hdr[c]).trim().toLowerCase();
        if (h === 'skill')       sCols.skill = c;
        if (h === 'ability' || h === 'attribute') sCols.ability = c;
        if (h === 'class skill') sCols.classes = c;
      }
      if (!('skill' in sCols) || !('ability' in sCols)) continue;

      for (var j = i + 1; j < data.length; j++) {
        var skillName = String(data[j][sCols.skill]).trim();
        if (!skillName) break;

        var classesRaw = ('classes' in sCols) ? String(data[j][sCols.classes] || '').trim() : '';
        var classesArr = classesRaw
          ? classesRaw.split(',').map(function(c) { return c.trim(); }).filter(function(c) { return c; })
          : [];

        skillTable.push({
          skill: skillName.toLowerCase(),
          ability: String(data[j][sCols.ability]).trim().toUpperCase(),
          classes: classesArr,
        });
      }
      break;
    }
  }
  return skillTable;
}


function findSkillRef(skillName, skillTable) {
  var name = skillName.toLowerCase().trim();

  for (var i = 0; i < skillTable.length; i++) {
    if (skillTable[i].skill === name) return skillTable[i];
  }

  for (var i = 0; i < skillTable.length; i++) {
    if (skillTable[i].skill.indexOf(name) === 0 || name.indexOf(skillTable[i].skill) === 0) {
      return skillTable[i];
    }
  }

  for (var i = 0; i < skillTable.length; i++) {
    if (skillTable[i].skill.indexOf(name) !== -1 || name.indexOf(skillTable[i].skill) !== -1) {
      return skillTable[i];
    }
  }

  return null;
}

function writeSkillSheet(ss, rows) {
  var sheetName = 'Skills';
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    var existingFilter = sheet.getFilter();
    if (existingFilter) existingFilter.remove();
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  if (rows.length <= 1) return;

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  // --- Replace sum column with formulas ---
  // sum = ability_score + ranks + class_mod + gm_mod
  // Columns: F=ability_score, G=ranks, H=class_mod, I=gm_mod, J=sum
  for (var r = 2; r <= rows.length; r++) {
    sheet.getRange(r, 10).setFormula('=F' + r + '+G' + r + '+H' + r + '+I' + r);
  }

  // Format header
  var headerRange = sheet.getRange(1, 1, 1, rows[0].length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1B1B2F');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  // Highlight gm_mod column to show it's editable
  var gmRange = sheet.getRange(2, 9, rows.length - 1, 1);
  gmRange.setBackground('#FFF9E6'); // yellow like input fields

  // Alternate row colors (skip gm_mod column)
  for (var r = 2; r <= rows.length; r++) {
    var bg = (r % 2 === 0) ? '#F3E5F5' : '#FFFFFF';
    sheet.getRange(r, 1, 1, 8).setBackground(bg);   // before gm_mod
    sheet.getRange(r, 10, 1, 1).setBackground(bg);   // sum column
  }

  // Column widths
  for (var c = 1; c <= rows[0].length; c++) {
    sheet.autoResizeColumn(c);
    var width = sheet.getColumnWidth(c);
    sheet.setColumnWidth(c, width + 30);
  }

  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, rows[0].length).createFilter();
  }

  sheet.setFrozenRows(1);
  sheet.setTabColor('#4A0E4E');
}