// ============================================================
// PF1e Character Database — Form Rules
// ============================================================
// Keeps the four Google Forms in line with what the scripts and
// Database formulas can parse:
//   applyFormValidation() — response validation on free-text questions
//   syncFormChoices()     — dropdown / checkbox options taken from the sheets
//
// Forms are found through their linked response sheets
// (sheet.getFormUrl()), so no form IDs are stored here.
// Questions are found by title (trimmed). A question that is missing
// or has an unexpected type is reported and skipped, never created:
// a new question writes to a new response column and breaks
// Column matchups and the Database formulas. Change question types
// in the form editor instead — that keeps the response column.
//
// Patterns are RE2 (Google's regex engine): no lookahead, no
// backreferences. Every pattern is anchored with ^...$.
// ============================================================


var FORM_SHEETS = {
  registration: 'Registration',
  lvlup:        'LvlUP',
  inventory:    'Inventory(raw)',
  games:        'Games',
};


// --- Building blocks ---

// Knowledge skills are written without parentheses ("knowledge arcana"),
// as in Reference tables; other skills may carry one: "perform (dance)".
// No other skill starts with K, so the second branch excludes it.
var RX_SKILL_NAME = "(?:[Kk]nowledge [A-Za-z]+|[A-JL-Za-jl-z][A-Za-z' ]*(?: ?\\([^(),]+\\))?)";

// Item name must not start with a digit: "Bedroll 2 0,1 5" splits into
// "Bedroll 2 0" and "1 5", and the second part has no name.
var RX_ITEM      = '[^,;\\d\\s][^,;]*';
var RX_LOOT_ITEM = '-?[^,;:\\d\\s-][^,;:]*';

var RX_AMOUNT = '[+-]?\\d+(?:\\.\\d+)?';

var RX = {
  // Names are keys and appear in "Name: …; Name: …" and comma lists
  charName: '^[^,;:]+$',
  vkLink:   '^(?:https?://)?(?:www\\.|m\\.)?vk\\.(?:com|ru)/[A-Za-z0-9_.]+/?$',
  ability:  '^(?:[3-9]|1[0-9]|20)$',
  level:    '^(?:[2-9]|1[0-9]|20)$',
  // "known/prepared", "-" means unlimited
  spellNum: '^\\s*(?:\\d+|-)\\s*/\\s*(?:\\d+|-)\\s*$',
  // Comma-separated list: ";" would glue entries together
  list:     '^[^;]+$',
  // Commas only between entries, none inside (...) or [...]
  entries:  '^(?:[^()\\[\\]]|\\([^(),]*\\)|\\[[^\\[\\],]*\\])*$',
  skills:    '^\\s*' + RX_SKILL_NAME + ' \\d+(?:\\s*,\\s*' + RX_SKILL_NAME + ' \\d+)*\\s*,?\\s*$',
  // Level-up: a skill without a number means +1 rank
  newSkills: '^\\s*' + RX_SKILL_NAME + '(?: \\d+)?(?:\\s*,\\s*' + RX_SKILL_NAME + '(?: \\d+)?)*\\s*,?\\s*$',
  items:    '^\\s*' + RX_ITEM + '(?:\\s*,\\s*' + RX_ITEM + ')*\\s*,?\\s*$',
  // Single-number answers are parsed by the sheet in its locale (ru_RU),
  // so a decimal comma is fine here. Inside lists only a dot works.
  gold:       '^\\d+(?:[.,]\\d+)?$',
  goldSigned: '^[+-]?\\d+(?:[.,]\\d+)?$',
  // "Name: -10; Name: +13"
  goldDelta: '^\\s*[^,;:]+:\\s*' + RX_AMOUNT + '(?:\\s*;\\s*[^,;:]+:\\s*' + RX_AMOUNT + ')*\\s*;?\\s*$',
  // "Name: item 1 30, item; Name: -item"
  loot: '^\\s*[^,;:]+:\\s*' + RX_LOOT_ITEM + '(?:\\s*,\\s*' + RX_LOOT_ITEM + ')*' +
        '(?:\\s*;\\s*[^,;:]+:\\s*' + RX_LOOT_ITEM + '(?:\\s*,\\s*' + RX_LOOT_ITEM + ')*)*\\s*;?\\s*$',
  // "Old Spell -> New Spell, Old -> New"
  retrain: '^\\s*[^,>]+->[^,>]+(?:\\s*,\\s*[^,>]+->[^,>]+)*\\s*$',
};

var HELP = {
  charName:  'Без запятых, точек с запятой и двоеточий',
  vkLink:    'Ссылка на страницу ВК: vk.com/имя, vk.ru/имя или https://vk.com/id12345',
  ability:   'Целое число от 3 до 20',
  level:     'Целое число от 2 до 20',
  whole:     'Целое число',
  spellNum:  'число известных/число подготовленных, "-" для бесконечности. Например: 5/- или 3/3',
  list:      'Через запятую, без точек с запятой',
  entries:   'Через запятую. Внутри скобок (...) и [...] запятых быть не должно',
  skills:    'Через запятую: навык и число рангов. Knowledge без скобок: knowledge arcana 1. Остальные со скобками: perform (dance) 2',
  newSkills: 'Через запятую: навык и число рангов (без числа = +1). Knowledge без скобок: knowledge arcana 1',
  items:     'Через запятую: название кол-во цена вес. Дробные — через точку: 0.5',
  gold:      'Число: 42 или 42,37',
  goldSigned: 'Число со знаком: -10 или +22,5',
  goldDelta: 'Имя: сумма; Имя: сумма. Дробные — через точку. Например: Аурелия: -10; Ке`цаль: +13.5',
  loot:      'Имя: предмет кол-во цена вес, предмет; Имя: -предмет. Дробные — через точку',
  retrain:   'Через запятую: Старое -> Новое',
  favClass:  'Название класса из списка, несколько — через /. Например: Bard или Fighter/Rogue',
};


// --- Static rules: what each free-text question accepts ---

// "Spells # rank known" -> ["Spells 1 rank known", …, "Spells 9 rank known"]
function spellLevelTitles(pattern) {
  var t = [];
  for (var i = 1; i <= 9; i++) t.push(pattern.replace(/#/g, i));
  return t;
}

var FORM_RULES = {
  registration: [
    { titles: ['Character name (short)'], rx: 'charName' },
    { titles: ['Vk link'], rx: 'vkLink' },
    { titles: ['Age', 'Height', 'Weight'], whole: true },
    { titles: ['Languages'], rx: 'list' },
    { titles: ['Archetype (if any)', 'Archetype second class (if any)', 'Archetype third class (if any)'], rx: 'entries' },
    { titles: ['STR', 'DEX', 'CON', 'INT', 'WIS', 'CHA'], rx: 'ability' },
    { titles: ['List your trained skills and their total ranks'], rx: 'skills' },
    { titles: ['List your feats', 'Class features', 'Special abilities', 'Traits', 'Drawbacks (if any)'], rx: 'entries' },
    { titles: ['Number of cantrips known'].concat(spellLevelTitles('Known spells#/Prepared spells#')), rx: 'spellNum' },
    { titles: ['Cantrips known'].concat(spellLevelTitles('Spells # rank known')), rx: 'list' },
    { titles: ['Remaining gold'], rx: 'gold' },
  ],
  lvlup: [
    { titles: ['Which level do you get'], rx: 'level' },
    { titles: ['HP rolled'], whole: true },
    { titles: ['New archetype', 'New feats', 'New special abilities', 'New class features', 'New Traits', 'New Drawbacks'], rx: 'entries' },
    { titles: ['New skill ranks'], rx: 'newSkills' },
    { titles: ['New cantrips num'].concat(spellLevelTitles('New spell # num')), rx: 'spellNum' },
    { titles: ['New cantrips'].concat(spellLevelTitles('New spells #')), rx: 'list' },
    { titles: ['Spells retrained'], rx: 'retrain' },
  ],
  inventory: [
    { titles: ['Items bought', 'Items sold', 'Items gifted', 'Items got'], rx: 'items' },
    { titles: ['Other gold changes'], rx: 'goldSigned' },
  ],
  games: [
    { titles: ['Party experience'], whole: true },
    { titles: ['Party gold reward'], rx: 'gold' },
    { titles: ['Other gold changes'], rx: 'goldDelta' },
    { titles: ['Additional rewards'], rx: 'loot' },
  ],
};


// ============================================================
// MAIN ENTRY POINTS
// ============================================================

function applyFormValidation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var report = newFormReport();

  for (var key in FORM_RULES) {
    var form = openLinkedForm(ss, key, report);
    if (!form) continue;
    var items = indexFormItems(form);

    var rules = FORM_RULES[key];
    for (var r = 0; r < rules.length; r++) {
      for (var t = 0; t < rules[r].titles.length; t++) {
        var title = rules[r].titles[t];
        if (rules[r].whole) {
          setWholeNumber(items, key, title, report);
        } else {
          setPattern(items, key, title, RX[rules[r].rx], HELP[rules[r].rx], report);
        }
      }
    }
  }

  showFormReport('Form validation applied', report);
}


function syncFormChoices() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var report = newFormReport();

  var ref = ss.getSheetByName('Reference tables');
  var classes = readRefColumn(ref, 'Class');
  var races   = readRefColumn(ref, 'Race');
  var deities = readRefColumn(ref, 'Deity');
  var chars   = readCharacterNames(ss);

  var favClassRx = '^\\s*(?:' + classes.map(escapeRegex).join('|') + ')' +
                   '(?:\\s*/\\s*(?:' + classes.map(escapeRegex).join('|') + '))*\\s*$';

  var reg = openLinkedForm(ss, 'registration', report);
  if (reg) {
    var regItems = indexFormItems(reg);
    setChoices(regItems, 'registration', 'Race', races, report);
    setChoices(regItems, 'registration', 'Class name', classes, report);
    setChoices(regItems, 'registration', 'Second class name', classes, report);
    setChoices(regItems, 'registration', 'Third class name', classes, report);
    setChoices(regItems, 'registration', 'Deity', deities, report);
    setPattern(regItems, 'registration', 'Favored class', favClassRx, HELP.favClass, report);
  }

  var lvl = openLinkedForm(ss, 'lvlup', report);
  if (lvl) {
    var lvlItems = indexFormItems(lvl);
    setChoices(lvlItems, 'lvlup', 'Character name', chars, report);
    setChoices(lvlItems, 'lvlup', 'Class to level up', classes, report);
    setPattern(lvlItems, 'lvlup', 'New Favorite Class', favClassRx, HELP.favClass, report);
  }

  var inv = openLinkedForm(ss, 'inventory', report);
  if (inv) setChoices(indexFormItems(inv), 'inventory', 'Character name', chars, report);

  var games = openLinkedForm(ss, 'games', report);
  if (games) setChoices(indexFormItems(games), 'games', 'Characters', chars, report);

  showFormReport('Form choices synced', report);
}


// ============================================================
// FORM HELPERS
// ============================================================

function openLinkedForm(ss, key, report) {
  var sheetName = FORM_SHEETS[key];
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) { report.errors.push('Sheet "' + sheetName + '" not found'); return null; }

  var url = sheet.getFormUrl();
  if (!url) {
    report.errors.push('Sheet "' + sheetName + '" is not linked to a form');
    return null;
  }
  return FormApp.openByUrl(url);
}

// Reads the form's questions once: { title: { item, type } }.
// Every getItems / getType / getTitle is a separate call to the Forms
// service, so looking a question up by scanning the form each time
// costs thousands of calls per run and takes minutes.
//
// Section headers can share a title with the question under them
// ("Race", "Traits"), so non-question items are skipped.
// The list is built here, not globally: global code also runs in
// onOpen, a simple trigger that is not allowed to touch FormApp.
function indexFormItems(form) {
  var nonQuestion = [
    FormApp.ItemType.PAGE_BREAK,
    FormApp.ItemType.SECTION_HEADER,
    FormApp.ItemType.IMAGE,
    FormApp.ItemType.VIDEO,
  ];
  var index = {};
  var items = form.getItems();
  for (var i = 0; i < items.length; i++) {
    var type = items[i].getType();
    if (nonQuestion.indexOf(type) !== -1) continue;
    var title = String(items[i].getTitle()).trim();
    if (!(title in index)) index[title] = { item: items[i], type: type };
  }
  return index;
}

function setPattern(items, key, title, pattern, help, report) {
  var entry = items[title];
  if (!entry) { report.missing.push(key + ': "' + title + '"'); return; }

  var item = entry.item, type = entry.type;
  if (type === FormApp.ItemType.TEXT) {
    item.asTextItem().setValidation(FormApp.createTextValidation()
      .setHelpText(help).requireTextMatchesPattern(pattern).build());
  } else if (type === FormApp.ItemType.PARAGRAPH_TEXT) {
    item.asParagraphTextItem().setValidation(FormApp.createParagraphTextValidation()
      .setHelpText(help).requireTextMatchesPattern(pattern).build());
  } else {
    report.wrongType.push(key + ': "' + title + '" is ' + type + ', expected text');
    return;
  }
  report.done.push(key + ': "' + title + '"');
}

function setWholeNumber(items, key, title, report) {
  var entry = items[title];
  if (!entry) { report.missing.push(key + ': "' + title + '"'); return; }
  if (entry.type !== FormApp.ItemType.TEXT) {
    report.wrongType.push(key + ': "' + title + '" is ' + entry.type + ', expected short text');
    return;
  }
  entry.item.asTextItem().setValidation(FormApp.createTextValidation()
    .setHelpText(HELP.whole).requireWholeNumber().build());
  report.done.push(key + ': "' + title + '"');
}

function setChoices(items, key, title, values, report) {
  var entry = items[title];
  if (!entry) { report.missing.push(key + ': "' + title + '"'); return; }
  if (values.length === 0) {
    report.errors.push(key + ': "' + title + '" — source list is empty, left unchanged');
    return;
  }

  var item = entry.item, type = entry.type;
  if (type === FormApp.ItemType.LIST) {
    item.asListItem().setChoiceValues(values);
  } else if (type === FormApp.ItemType.CHECKBOX) {
    item.asCheckboxItem().setChoiceValues(values);
  } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
    item.asMultipleChoiceItem().setChoiceValues(values);
  } else {
    report.wrongType.push(key + ': "' + title + '" is ' + type +
      ' — change it to a dropdown/checkboxes in the form editor');
    return;
  }
  report.done.push(key + ': "' + title + '" (' + values.length + ' options)');
}


// ============================================================
// DATA SOURCES
// ============================================================

// Reference tables keep column headers in row 2; values run down
// until the first empty cell.
function readRefColumn(ref, header) {
  var data = ref.getDataRange().getValues();
  var col = data[1].map(function(h) { return String(h).trim(); }).indexOf(header);
  if (col === -1) throw new Error('Column "' + header + '" not found in Reference tables row 2');

  var values = [];
  for (var r = 2; r < data.length; r++) {
    var v = String(data[r][col]).trim();
    if (!v) break;
    values.push(v);
  }
  return values;
}

function readCharacterNames(ss) {
  var data = ss.getSheetByName('Database').getDataRange().getValues();
  var col = data[3].map(function(h) { return String(h).trim(); }).indexOf('char');
  var names = [];
  for (var r = 4; r < data.length; r++) {
    var v = String(data[r][col]).trim();
    if (v && names.indexOf(v) === -1) names.push(v);
  }
  return names.sort();
}

function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\\/]/g, '\\$&');
}


// ============================================================
// REPORT
// ============================================================

function newFormReport() {
  return { done: [], missing: [], wrongType: [], errors: [] };
}

function showFormReport(header, report) {
  var msg = header + '\n\n✅ Updated: ' + report.done.length + '\n';
  if (report.missing.length > 0) {
    msg += '\n⚠️ Question not found (title changed?):\n  • ' + report.missing.join('\n  • ') + '\n';
  }
  if (report.wrongType.length > 0) {
    msg += '\n⚠️ Unexpected question type:\n  • ' + report.wrongType.join('\n  • ') + '\n';
  }
  if (report.errors.length > 0) {
    msg += '\n❌ Errors:\n  • ' + report.errors.join('\n  • ') + '\n';
  }
  Logger.log(msg + '\nDetails:\n  ' + report.done.join('\n  '));
  SpreadsheetApp.getUi().alert(msg);
}
