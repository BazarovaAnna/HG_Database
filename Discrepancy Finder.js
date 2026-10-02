// ============================================================
// PF1e Character Database — Discrepancy Checker v2
// ============================================================
// Validates character data against PF1e rules and flags issues.
//
// SEVERITY LEVELS:
//   ❌ ERROR   — definitely wrong, rules violation
//   ⚠️ WARNING — likely wrong, needs GM review
//   ℹ️ MINOR   — cosmetic or soft rule, worth a glance
//
// SETUP:
//   Add this to your existing Apps Script project.
//   The onOpen menu already has 'Find the data Discrepancies' → findDiscrepancies
//
// REQUIRES:
//   - "Database" sheet with row 4 as field keys
//   - "Reference tables" sheet with Class table (BAB Type, Good Fort/Ref/Will)
//     and BAB & Save Progression table (Level, Full, 3/4, 1/2, Good Save, Bad Save)
//   - Optionally: Deity table on Reference tables (Deity | Alignment columns)
// ============================================================


// --- REFERENCE DATA BUILT INTO SCRIPT ---

var ALIGNMENT_COORDS = {
  'Lawful Good':     [0, 0], 'Neutral Good':  [1, 0], 'Chaotic Good':    [2, 0],
  'Lawful Neutral':  [0, 1], 'True Neutral':  [1, 1], 'Neutral':         [1, 1],
  'Chaotic Neutral': [2, 1],
  'Lawful Evil':     [0, 2], 'Neutral Evil':  [1, 2], 'Chaotic Evil':    [2, 2],
};

var ALIGNMENT_RESTRICTIONS = {
  'Paladin':   { must: ['Lawful Good'], label: 'must be Lawful Good' },
  'Monk':      { must: ['Lawful Good','Lawful Neutral','Lawful Evil'], label: 'must be Lawful' },
  'Barbarian': { forbidden: ['Lawful Good','Lawful Neutral','Lawful Evil'], label: 'cannot be Lawful' },
};

var CLASS_BONUS_FEAT_MILESTONES = {
  'Fighter':    [1,2,4,6,8,10,12,14,16,18,20],
  'Monk':       [1,2,6,10,14,18],
  'Ranger':     [2,6,10,14,18],
  'Wizard':     [5,10,15,20],
  'Bloodrager': [6,9,12,15,18],
};


// ============================================================
// MAIN ENTRY POINT
// ============================================================

function findDiscrepancies() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var main = ss.getSheetByName('Database');
  var ref = ss.getSheetByName('Reference tables');
  if (!main) { SpreadsheetApp.getUi().alert('Sheet "Database" not found!'); return; }
  if (!ref) { SpreadsheetApp.getUi().alert('Sheet "Reference tables" not found!'); return; }

  var data = main.getDataRange().getValues();
  var headers = data[3]; // row 4 = index 3

  // --- Build column lookup ---
  var col = {};
  for (var c = 0; c < headers.length; c++) {
    var key = String(headers[c]).trim();
    if (!key) continue;
    if (!(key in col)) { col[key] = c; }
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

  // --- Load reference data ---
  var refData = loadRefData(ref);
  var classTable = refData.classes;
  var babProg = refData.bab;
  var saveProg = refData.saves;
  var deityTableFromRef = refData.deities;

  // --- Run checks ---
  var issues = [];
  issues.push(['Персонаж', 'Severity', 'Category', 'Issue', 'Expected', 'Actual']);

  for (var i = 4; i < data.length; i++) {
    var row = data[i];
    var charName = String(v(row, 'char')).trim();
    if (!charName) continue;

    // Gather common values
    var race      = String(v(row, 'race')).trim();
    var variant   = String(v(row, 'variant')).trim();
    var alignment = String(v(row, 'alignment')).trim();
    var deity     = String(v(row, 'deity')).trim();

    var class1     = String(v(row, 'class_1')).trim();
    var class1Lvl  = vNum(row, 'class_1_lvl');
    var class2     = String(v(row, 'class_2')).trim();
    var class2Lvl  = vNum(row, 'class_2_lvl');
    var class3     = String(v(row, 'class_3')).trim();
    var class3Lvl  = vNum(row, 'class_3_lvl');

    var inputLvl  = vNum(row, 'input_lvl');
    var level     = vNum(row, 'level');
    var favClass  = String(v(row, 'favored_class')).trim();

    var fcHP    = vNum(row, 'favored_class_bonus_HP');
    var fcSkill = vNum(row, 'favored_class_bonus_skill');
    var fcRace  = vNum(row, 'favored_class_bonus_race');
    var fcTotal = vNum(row, 'favored_class_bonus');

    var strScore = vNum(row, 'str_score');
    var dexScore = vNum(row, 'dex_score');
    var conScore = vNum(row, 'con_score');
    var intScore = vNum(row, 'int_score');
    var wisScore = vNum(row, 'wis_score');
    var chaScore = vNum(row, 'cha_score');

    var strMod = vNum(row, 'str_mod');
    var dexMod = vNum(row, 'dex_mod');
    var conMod = vNum(row, 'con_mod');
    var intMod = vNum(row, 'int_mod');
    var wisMod = vNum(row, 'wis_mod');
    var chaMod = vNum(row, 'cha_mod');

    var maxHP   = vNum(row, 'max_hp');
    var minHP   = vNum(row, 'min_hp');
    var hpTotal = vNum(row, 'hp_total');

    var bab = vNum(row, 'bab');

    var skillPtsUsed = vNum(row, 'skill_pts_used');
    var skillPoints  = vNum(row, 'skill_points');

    var classes = [];
    if (class1) classes.push({ name: class1, lvl: class1Lvl });
    if (class2) classes.push({ name: class2, lvl: class2Lvl });
    if (class3) classes.push({ name: class3, lvl: class3Lvl });


    // ========================================
    // CHECK 1: Input level vs calculated level
    // ========================================
    if (inputLvl > 0 && level > 0 && inputLvl !== level) {
      issues.push([charName, '❌ ERROR', 'Level',
        'Input level does not match sum of class levels',
        'input_lvl=' + inputLvl, 'calculated=' + level]);
    }


    // ========================================
    // CHECK 2: Alignment vs Deity (within 1 step)
    // ========================================
    if (deity && alignment) {
      var deityAlign = deityTableFromRef[deity] || null;
      if (deityAlign) {
        var dist = alignmentDistance(alignment, deityAlign);
        if (dist > 1) {
          issues.push([charName, 'ℹ️ MINOR', 'Alignment',
            'Alignment is ' + dist + ' steps from deity ' + deity + ' (' + deityAlign + ')',
            'Within 1 step', alignment + ' vs ' + deityAlign]);
        }
      }
    }


    // ========================================
    // CHECK 3: Class alignment restrictions
    // ========================================
    for (var ci = 0; ci < classes.length; ci++) {
      var cls = classes[ci].name;
      if (cls in ALIGNMENT_RESTRICTIONS) {
        var restriction = ALIGNMENT_RESTRICTIONS[cls];
        if (restriction.must && restriction.must.indexOf(alignment) === -1) {
          issues.push([charName, '❌ ERROR', 'Alignment',
            cls + ' ' + restriction.label,
            restriction.must.join(' or '), alignment]);
        }
        if (restriction.forbidden && restriction.forbidden.indexOf(alignment) !== -1) {
          issues.push([charName, '❌ ERROR', 'Alignment',
            cls + ' ' + restriction.label,
            'Not ' + restriction.forbidden.join('/'), alignment]);
        }
      }
    }


    // ========================================
    // CHECK 4: Favored class is one of the character's classes
    // ========================================
    var favClasses = favClass.split('/').map(function(f) { return f.trim().toLowerCase(); });
    var classNamesLower = classes.map(function(c) { return c.name.toLowerCase(); });
    for (var fi = 0; fi < favClasses.length; fi++) {
      if (classNamesLower.indexOf(favClasses[fi]) === -1) {
        issues.push([charName, '❌ ERROR', 'Favored Class',
          'Favored class "' + favClasses[fi] + '" is not one of the character\'s classes',
        classNamesLower.join(', '), favClasses[fi]]);
      }
    }

    // ========================================
    // CHECK 5: Every class exists in the reference table
    // ========================================
    for (var ci = 0; ci < classes.length; ci++) {
      if (!classTable[classes[ci].name]) {
        issues.push([charName, '❌ ERROR', 'Class',
          'Class "' + classes[ci].name + '" not found in Reference Tables — BAB and save checks will be skipped',
          'Add to Reference Tables or fix spelling', classes[ci].name]);
      }
    }


    // ========================================
    // CHECK 6: Favored class bonus count matches favored class level
    // ========================================
    if (favClass) {
      var favClassLvl = 0;
      // Handle multiple favored classes (Half-Elf)
      var favClasses = favClass.split('/').map(function(f) { return f.trim().toLowerCase(); });
      var favClassLvl = 0;
      for (var ci = 0; ci < classes.length; ci++) {
        if (favClasses.indexOf(classes[ci].name.toLowerCase()) !== -1) {
          favClassLvl += classes[ci].lvl;
        }
      }

      var bonusSum = fcHP + fcSkill + fcRace;
      if (bonusSum !== favClassLvl) {
        issues.push([charName, '⚠️ WARNING', 'Favored Class',
          'FC bonus total (' + bonusSum + ') ≠ favored class level (' + favClassLvl + ')',
          favClassLvl + ' bonuses', bonusSum + ' (HP:' + fcHP + ' Skill:' + fcSkill + ' Race:' + fcRace + ')']);
      }
    }


    // ========================================
    // CHECK 7: Raw ability scores > 20
    // ========================================
    var abilityNames = ['str_score', 'dex_score', 'con_score', 'int_score', 'wis_score', 'cha_score'];
    var abilityLabels = ['STR', 'DEX', 'CON', 'INT', 'WIS', 'CHA'];
    for (var ai = 0; ai < abilityNames.length; ai++) {
      var rawScore = vNum(row, abilityNames[ai]);
      if (rawScore > 20) {
        issues.push([charName, '⚠️ WARNING', 'Abilities',
          abilityLabels[ai] + ' raw score is ' + rawScore + ' (above 20)',
          '≤ 20 for point buy', rawScore]);
      }
    }


    // ========================================
    // CHECK 8: Racial variant filled but racial_var columns EMPTY
    //          (0 is valid — empty/blank is the problem)
    // ========================================
    if (variant) {
      var racialVarKeys = [
        { key: 'str_racial_var', label: 'STR' },
        { key: 'dex_racial_var', label: 'DEX' },  
        { key: 'con_racial_var', label: 'CON' },
        { key: 'int_racial_var', label: 'INT' },
        { key: 'wis_racial_var', label: 'WIS' },
        { key: 'cha_racial_var', label: 'CHA' },
      ];
      var emptyRacialVars = [];
      for (var rv = 0; rv < racialVarKeys.length; rv++) {
        var rvKey = racialVarKeys[rv].key;
        if (!(rvKey in col)) continue; // column doesn't exist, skip
        var rvVal = row[col[rvKey]];
        if (rvVal === '' || rvVal === undefined || rvVal === null) {
          emptyRacialVars.push(racialVarKeys[rv].label);
        }
      }
      if (emptyRacialVars.length > 0) {
        issues.push([charName, '⚠️ WARNING', 'Abilities',
          'Racial variant "' + variant + '" is set but racial_var is empty for: ' + emptyRacialVars.join(', '),
          'Fill all racial_var columns (0 is fine, empty is not)', emptyRacialVars.length + ' empty']);
      }
    }


    // ========================================
    // CHECK 9: HP total vs min/max HP
    // ========================================
    if (hpTotal > 0) {
      if (maxHP > 0 && hpTotal > maxHP) {
        issues.push([charName, '❌ ERROR', 'HP',
          'HP (' + hpTotal + ') exceeds maximum possible (' + maxHP + ')',
          '≤ ' + maxHP, hpTotal]);
      }
      if (minHP > 0 && hpTotal < minHP) {
        issues.push([charName, '❌ ERROR', 'HP',
          'HP (' + hpTotal + ') is below minimum possible (' + minHP + ')',
          '≥ ' + minHP, hpTotal]);
      }
      if (maxHP > 0 && hpTotal === maxHP && level > 1) {
        issues.push([charName, 'ℹ️ MINOR', 'HP',
          'HP exactly equals maximum (' + maxHP + ') — all max rolls?',
          'Verify rolling method', hpTotal]);
      }
      if (minHP > 0 && hpTotal === minHP && level > 1) {
        issues.push([charName, '⚠️ WARNING', 'HP',
          'HP exactly equals minimum (' + minHP + ') — all 1s rolled?',
          'Extremely unlikely', hpTotal]);
      }
    }


    // ========================================
    // CHECK 10: BAB from ref table (multiclass)
    // ========================================
    if (classes.length > 0 && bab >= 0) {
      var expectedBAB = 0;
      var babCheckable = true;
      for (var ci = 0; ci < classes.length; ci++) {
        var clsInfo = classTable[classes[ci].name];
        if (!clsInfo || !babProg[clsInfo.babType]) { babCheckable = false; break; }
        var lookup = babProg[clsInfo.babType][classes[ci].lvl];
        if (lookup === undefined) { babCheckable = false; break; }
        expectedBAB += lookup;
      }
      if (babCheckable && expectedBAB !== bab) {
        issues.push([charName, '❌ ERROR', 'BAB',
          'BAB mismatch',
          'Expected ' + expectedBAB + ' from ref table (' +
            classes.map(function(c) { return c.name + ' ' + c.lvl; }).join(' / ') + ')',
          'Got ' + bab]);
      }
    }


    // ========================================
    // CHECK 11: Save base values from ref table (multiclass)
    // ========================================
    var saveChecks = [
      { baseName: 'fort_base', refKey: 'goodFort', label: 'Fortitude' },
      { baseName: 'ref_base',  refKey: 'goodRef',  label: 'Reflex' },
      { baseName: 'will_base', refKey: 'goodWill', label: 'Will' },
    ];
    for (var si = 0; si < saveChecks.length; si++) {
      var sc = saveChecks[si];
      if (!(sc.baseName in col)) continue;

      var actualBase = vNum(row, sc.baseName);
      var expectedBase = 0;
      var saveCheckable = true;
      for (var ci = 0; ci < classes.length; ci++) {
        var clsInfo = classTable[classes[ci].name];
        if (!clsInfo) { saveCheckable = false; break; }
        var isGood = (clsInfo[sc.refKey] === 'Yes');
        var progKey = isGood ? 'good' : 'bad';
        if (!saveProg[progKey] || saveProg[progKey][classes[ci].lvl] === undefined) {
          saveCheckable = false; break;
        }
        expectedBase += saveProg[progKey][classes[ci].lvl];
      }
      if (saveCheckable && expectedBase !== actualBase) {
        issues.push([charName, '❌ ERROR', 'Saves',
          sc.label + ' base save mismatch',
          'Expected ' + expectedBase + ' from ref table (' +
            classes.map(function(c) {
              var ci2 = classTable[c.name];
              var good = ci2 ? (ci2[sc.refKey] === 'Yes' ? 'good' : 'bad') : '?';
              return c.name + ' ' + c.lvl + ' (' + good + ')';
            }).join(' + ') + ')',
          'Got ' + actualBase]);
      }
    }

    // Also check save totals if base columns are missing but totals exist
    var saveTotalChecks = [
      { totalName: 'fort_total', baseName: 'fort_base', miscName: 'fort_misc', refKey: 'goodFort', modVal: conMod, label: 'Fortitude' },
      { totalName: 'ref_total',  baseName: 'ref_base',  miscName: 'ref_misc',  refKey: 'goodRef',  modVal: dexMod, label: 'Reflex' },
      { totalName: 'will_total', baseName: 'will_base', miscName: 'will_misc', refKey: 'goodWill', modVal: wisMod, label: 'Will' },
    ];
    for (var si = 0; si < saveTotalChecks.length; si++) {
      var sc = saveTotalChecks[si];
      // Only check total if we did NOT already check the base (avoid double-reporting)
      if (sc.baseName in col) continue;
      if (!(sc.totalName in col)) continue;

      var actualTotal = vNum(row, sc.totalName);
      var expectedBase = 0;
      var saveCheckable = true;
      for (var ci = 0; ci < classes.length; ci++) {
        var clsInfo = classTable[classes[ci].name];
        if (!clsInfo) { saveCheckable = false; break; }
        var isGood = (clsInfo[sc.refKey] === 'Yes');
        var progKey = isGood ? 'good' : 'bad';
        if (!saveProg[progKey] || saveProg[progKey][classes[ci].lvl] === undefined) {
          saveCheckable = false; break;
        }
        expectedBase += saveProg[progKey][classes[ci].lvl];
      }
      if (!saveCheckable) continue;
      var misc = (sc.miscName in col) ? vNum(row, sc.miscName) : 0;
      var expectedTotal = expectedBase + sc.modVal + misc;
      if (expectedTotal !== actualTotal) {
        issues.push([charName, '⚠️ WARNING', 'Saves',
          sc.label + ' total mismatch (base ' + expectedBase + ' + mod ' + sc.modVal + ' + misc ' + misc + ')',
          'Expected ' + expectedTotal, 'Got ' + actualTotal]);
      }
    }


    // ========================================
    // CHECK 12: Feat count (warning only)
    // ========================================
    var featsStr = String(v(row, 'feats')).trim();
    if (featsStr) {
      var featCount = featsStr.split(',').filter(function(f) { return f.trim() !== ''; }).length;
      var expectedBaseFeats = Math.floor((level + 1) / 2);
      var humanBonus = (race === 'Human') ? 1 : 0;
      var classBonusFeats = 0;
      for (var ci = 0; ci < classes.length; ci++) {
        var milestones = CLASS_BONUS_FEAT_MILESTONES[classes[ci].name];
        if (milestones) {
          classBonusFeats += milestones.filter(function(m) { return classes[ci].lvl >= m; }).length;
        }
      }
      var expectedMinFeats = expectedBaseFeats + humanBonus + classBonusFeats;
      if (featCount < expectedMinFeats) {
        issues.push([charName, '⚠️ WARNING', 'Feats',
          'Character has ' + featCount + ' feats, expected at least ' + expectedMinFeats +
          ' (base:' + expectedBaseFeats + ' human:' + humanBonus + ' class:' + classBonusFeats + ')',
          '≥ ' + expectedMinFeats, featCount]);
      }
      if (featCount > expectedMinFeats + 3) {
        issues.push([charName, '⚠️ WARNING', 'Feats',
          'Character has ' + featCount + ' feats, expected ~' + expectedMinFeats +
          ' — too many? (extra from archetypes/items?)',
          '~' + expectedMinFeats, featCount]);
      }
    }


    // ========================================
    // CHECK 13: Class features / special abilities (basic count)
    // ========================================
    var classFeaturesStr = String(v(row, 'class_features')).trim();
    var specialAbilitiesStr = String(v(row, 'special_abilities')).trim();

    if (level >= 3 && (!classFeaturesStr || classFeaturesStr === '')) {
      issues.push([charName, '⚠️ WARNING', 'Features',
        'No class features listed for a level ' + level + ' character',
        'At least a few class features', 'Empty']);
    }

    if (classFeaturesStr) {
      var featureCount = classFeaturesStr.split(',').filter(function(f) { return f.trim() !== ''; }).length;
      if (featureCount < Math.floor(level / 3) && level >= 4) {
        issues.push([charName, 'ℹ️ MINOR', 'Features',
          'Only ' + featureCount + ' class features for level ' + level + ' — seems low',
          '≥ ' + Math.floor(level / 3) + ' features', featureCount]);
      }
    }


    // ========================================
    // CHECK 14: Skill points used vs available
    // ========================================
    if (skillPoints > 0) {
      if (skillPtsUsed > skillPoints + 2*level) {
        issues.push([charName, '❌ ERROR', 'Skills',
          'Spent ' + skillPtsUsed + ' skill points but only ' + skillPoints + ' available',
          skillPoints + ' max', skillPtsUsed + ' used']);
      } else if(skillPtsUsed > skillPoints) {
        issues.push([charName, '⚠️ WARNING', 'Skills',
          'Spent ' + skillPtsUsed + ' skill points but only ' + skillPoints + ' available',
          skillPoints + ' max', skillPtsUsed + ' used']);
      }
      if (skillPtsUsed < skillPoints) {
        issues.push([charName, '⚠️ WARNING', 'Skills',
          'Only spent ' + skillPtsUsed + '/' + skillPoints + ' skill points — ' +
          (skillPoints - skillPtsUsed) + ' unspent',
          skillPoints, skillPtsUsed]);
      }
    }
    // ========================================
    // CHECK 15: is_caster matches reference table
    // ========================================
    var isCasterInput = String(v(row, 'is_caster')).trim().toLowerCase();
    if (isCasterInput) {
      var shouldBeCaster = false;
      for (var ci = 0; ci < classes.length; ci++) {
        var clsInfo = classTable[classes[ci].name];
        if (clsInfo && clsInfo.caster === 'Yes') {
          shouldBeCaster = true;
          break;
        }
      }
      var claimsIsCaster = (isCasterInput === 'yes' || isCasterInput === 'да' || isCasterInput === 'true');

      if (claimsIsCaster && !shouldBeCaster) {
        issues.push([charName, '⚠️ WARNING', 'Spells',
          'Marked as caster but no casting class found in reference table',
          'Non-caster classes: ' + classes.map(function(c) { return c.name; }).join(', '),
          'is_caster = ' + isCasterInput]);
      }
      if (!claimsIsCaster && shouldBeCaster) {
        var casterClasses = classes.filter(function(c) {
          var ci2 = classTable[c.name];
          return ci2 && ci2.caster === 'Yes';
        }).map(function(c) { return c.name; });
        issues.push([charName, '⚠️ WARNING', 'Spells',
          'Not marked as caster but has casting class(es): ' + casterClasses.join(', '),
          'Should be is_caster = Yes', 'is_caster = ' + isCasterInput]);
      }
    }


    // ========================================
    // CHECK 16: Spell list count vs _num field
    // ========================================
    var spellLevels = [
      { numKey: 'cant_num', listKey: 'cant', label: 'Cantrips' }
    ];
    for (var sl = 1; sl <= 9; sl++) {
      spellLevels.push({
        numKey: 'spell' + sl + '_num',
        listKey: 'spell' + sl,
        label: 'Level ' + sl
      });
    }

    for (var si = 0; si < spellLevels.length; si++) {
      var sp = spellLevels[si];
      var numStr = String(v(row, sp.numKey)).trim();
      var listStr = String(v(row, sp.listKey)).trim();
      if (!numStr || numStr === '0/0') continue;

      var parts = numStr.split('/');
      var known = (parts[0].trim() === '-') ? -1 : (parseInt(parts[0]) || 0);
      var prepared = (parts[1] && parts[1].trim() === '-') ? -1 : (parseInt(parts[1]) || 0);
      var listed = listStr ? listStr.split(',').filter(function(s) { return s.trim(); }).length : 0;

      // Check: list should match either known or prepared
      if (known >= 0 && prepared >= 0 && listed !== known && listed !== prepared) {
        issues.push([charName, '⚠️ WARNING', 'Spells',
          sp.label + ': listed ' + listed + ' spells but num says ' + numStr,
          known + ' or ' + prepared, listed + ' listed']);
      }
      if (known === -1 && prepared >= 0 && listed !== prepared) {
        issues.push([charName, '⚠️ WARNING', 'Spells',
          sp.label + ': listed ' + listed + ' spells but should have ' + prepared + ' prepared',
          prepared + ' prepared', listed + ' listed']);
      }
      if (known >= 0 && prepared === -1 && listed !== known) {
        issues.push([charName, '⚠️ WARNING', 'Spells',
          sp.label + ': listed ' + listed + ' spells but should know ' + known,
          known + ' known', listed + ' listed']);
      }

      // Check: has _num but no spell list
      if ((known > 0 || prepared > 0) && listed === 0) {
        issues.push([charName, '⚠️ WARNING', 'Spells',
          sp.label + ': has ' + numStr + ' but spell list is empty',
          'List the spells', 'Empty']);
      }

      // Check: has spell list but _num is empty/missing
      if (listed > 0 && !numStr) {
        issues.push([charName, 'ℹ️ MINOR', 'Spells',
          sp.label + ': has ' + listed + ' spells listed but no _num value',
          'Fill in known/prepared count', listed + ' listed']);
      }
    }

  } // end character loop


  // --- Write results ---
  writeDiscrepancySheet(ss, issues);

  // --- Summary alert ---
  var errorCount = issues.filter(function(r) { return r[1] === '❌ ERROR'; }).length;
  var warnCount  = issues.filter(function(r) { return r[1] === '⚠️ WARNING'; }).length;
  var minorCount = issues.filter(function(r) { return r[1] === 'ℹ️ MINOR'; }).length;
  var total = errorCount + warnCount + minorCount;

  SpreadsheetApp.getUi().alert(
    'Discrepancy check complete!\n\n' +
    '❌ Errors: ' + errorCount + '\n' +
    '⚠️ Warnings: ' + warnCount + '\n' +
    'ℹ️ Minor: ' + minorCount + '\n' +
    '─────────────\n' +
    'Total: ' + total + ' issues\n\n' +
    'See the "Discrepancies" sheet for details.'
  );
}


// ============================================================
// ALIGNMENT HELPER
// ============================================================

function alignmentDistance(align1, align2) {
  var c1 = ALIGNMENT_COORDS[align1];
  var c2 = ALIGNMENT_COORDS[align2];
  if (!c1 || !c2) return -1;
  return Math.max(Math.abs(c1[0] - c2[0]), Math.abs(c1[1] - c2[1]));
}


// ============================================================
// REFERENCE TABLE LOADER (classes + BAB/save progressions + deities)
// ============================================================

function loadRefData(refSheet) {
  var data = refSheet.getDataRange().getValues();
  var classTable = {};
  var babProgression = {};
  var saveProgression = {};
  var deityTable = {};

  for (var i = 0; i < data.length; i++) {
    var rowStr = data[i].join('|');

    // --- Find CLASS table ---
    if (rowStr.indexOf('BAB') !== -1 && rowStr.indexOf('Class') !== -1) {
      var hdr = data[i];
      var colMap = {};
      for (var c = 0; c < hdr.length; c++) {
        var h = String(hdr[c]).trim();
        if (h === 'Class')                          colMap.name = c;
        if (h === 'BAB Type' || h === 'BAB')        colMap.babType = c;
        if (h === 'Good Fort' || h === 'Fort')      colMap.goodFort = c;
        if (h === 'Good Ref'  || h === 'Ref')       colMap.goodRef = c;
        if (h === 'Good Will' || h === 'Will')      colMap.goodWill = c;
        if (h === 'Skills/Lvl' || h === 'Skills')   colMap.skillsPerLvl = c;
        if (h === 'Hit Die')                        colMap.hitDie = c;
        if (h === 'Caster?' || h === 'Caster')      colMap.caster = c;
      }
      for (var j = i + 1; j < data.length; j++) {
        var className = String(data[j][colMap.name]).trim();
        if (!className) break;
        classTable[className] = {
          babType:      String(data[j][colMap.babType]).trim(),
          goodFort:     String(data[j][colMap.goodFort]).trim(),
          goodRef:      String(data[j][colMap.goodRef]).trim(),
          goodWill:     String(data[j][colMap.goodWill]).trim(),
          skillsPerLvl: Number(data[j][colMap.skillsPerLvl]) || 0,
          hitDie:       Number(data[j][colMap.hitDie]) || 0,
          caster: String(data[j][colMap.caster] || '').trim(),
        };
      }
    }

    // --- Find BAB & SAVE PROGRESSION table ---
    if (rowStr.indexOf('Level') !== -1 &&
        (rowStr.indexOf('Full') !== -1 || rowStr.indexOf('Good Save') !== -1)) {
      var hdr = data[i];
      var progCols = {};
      for (var c = 0; c < hdr.length; c++) {
        var h = String(hdr[c]).trim();
        if (h === 'Level')                              progCols.level = c;
        if (h === 'Full BAB' || h === 'Full')           progCols.fullBAB = c;
        if (h === '3/4 BAB'  || h === '3/4')            progCols.threeQuarterBAB = c;
        if (h === '1/2 BAB'  || h === '1/2')            progCols.halfBAB = c;
        if (h === 'Good Save' || h === 'Good')          progCols.goodSave = c;
        if (h === 'Bad Save'  || h === 'Bad')           progCols.badSave = c;
      }

      if ('level' in progCols) {
        for (var j = i + 1; j < data.length; j++) {
          var lvl = Number(data[j][progCols.level]);
          if (!lvl || lvl < 1) break;

          if ('fullBAB' in progCols) {
            if (!babProgression['Full']) babProgression['Full'] = {};
            babProgression['Full'][lvl] = Number(data[j][progCols.fullBAB]) || 0;
          }
          if ('threeQuarterBAB' in progCols) {
            if (!babProgression['3/4']) babProgression['3/4'] = {};
            babProgression['3/4'][lvl] = Number(data[j][progCols.threeQuarterBAB]) || 0;
          }
          if ('halfBAB' in progCols) {
            if (!babProgression['1/2']) babProgression['1/2'] = {};
            babProgression['1/2'][lvl] = Number(data[j][progCols.halfBAB]) || 0;
          }
          if ('goodSave' in progCols) {
            if (!saveProgression['good']) saveProgression['good'] = {};
            saveProgression['good'][lvl] = Number(data[j][progCols.goodSave]) || 0;
          }
          if ('badSave' in progCols) {
            if (!saveProgression['bad']) saveProgression['bad'] = {};
            saveProgression['bad'][lvl] = Number(data[j][progCols.badSave]) || 0;
          }
        }
      }
    }

    // --- Find DEITY table ---
    if (rowStr.indexOf('Deity') !== -1 && rowStr.indexOf('Alignment') !== -1) {
      var hdr = data[i];
      var dCol = -1, aCol = -1;
      for (var c = 0; c < hdr.length; c++) {
        if (String(hdr[c]).trim() === 'Deity') dCol = c;
        if (String(hdr[c]).trim() === 'Alignment') aCol = c;
      }
      if (dCol !== -1 && aCol !== -1) {
        for (var j = i + 1; j < data.length; j++) {
          var dName = String(data[j][dCol]).trim();
          var dAlign = String(data[j][aCol]).trim();
          if (!dName) break;
          deityTable[dName] = dAlign;
        }
      }
    }
  }

  return {
    classes: classTable,
    bab: babProgression,
    saves: saveProgression,
    deities: deityTable,
  };
}


// ============================================================
// OUTPUT WRITER
// ============================================================

function writeDiscrepancySheet(ss, rows) {
  var sheetName = 'Discrepancies';
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    var existingFilter = sheet.getFilter();
    if (existingFilter) existingFilter.remove();
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  if (rows.length === 0) return;

  var range = sheet.getRange(1, 1, rows.length, rows[0].length);
  range.setValues(rows);

  // Format header
  var headerRange = sheet.getRange(1, 1, 1, rows[0].length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1B1B2F');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  // Severity-based row coloring
  for (var r = 2; r <= rows.length; r++) {
    var severity = rows[r - 1][1];
    var bg = '#FFFFFF';
    if (severity === '❌ ERROR')    bg = '#FFCDD2';
    if (severity === '⚠️ WARNING') bg = '#FFF9C4';
    if (severity === 'ℹ️ MINOR')   bg = '#E3F2FD';
    sheet.getRange(r, 1, 1, rows[0].length).setBackground(bg);
  }

  // Auto-resize
  for (var c = 1; c <= rows[0].length; c++) {
    sheet.autoResizeColumn(c);
    var width = sheet.getColumnWidth(c);
    sheet.setColumnWidth(c, width + 30);
  }
  
  sheet.setColumnWidth(4, 400); // Issue
  sheet.setColumnWidth(5, 200); // Expected
  sheet.setColumnWidth(6, 200); // Actual

  // Add filter
  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, rows[0].length).createFilter();
  }

  sheet.setFrozenRows(1);
  sheet.setTabColor('#FF0000');
}