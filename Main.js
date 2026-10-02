// ============================================================
// MAIN ENTRY POINTS
// ============================================================

function parseAll() {
  parseFeats();
  parseSkills();
  parseSpells();
  SpreadsheetApp.getUi().alert('Feats, Skills, and Spells tables updated!');
}

function processTAll() {
  initializeStartingGold();
  processGameSessions();
  processInventory();
  SpreadsheetApp.getUi().alert('Transactions updated!');
}

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
    .addItem('Process ALL Transactions', 'processTAll')
    .addSeparator()
    .addItem('Process Starting Gold -> Transactions', 'initializeStartingGold')
    .addItem('Process Game Sessions -> Transactions', 'processGameSessions')
    .addItem('Process Inventory Form -> Transactions', 'processInventory')
    .addSeparator()
    .addItem('Rebuild Inventory Sheet', 'rebuildInventory')
    .addSeparator()
    .addItem('Find data Discrepancies', 'findDiscrepancies')
    .addSeparator()
    .addItem('Sync form choices', 'syncFormChoices')
    .addItem('Apply form validation', 'applyFormValidation')
    .addToUi();
}

// --- CONFIG ---
// Adjust these if your column layout changes
const CONFIG = {
  mainSheet:      'Database',    // name of the characters sheet
  featsSheet:     'Feats',       // will be created/overwritten
  skillsSheet:    'Skills',      // will be created/overwritten
  headerRow:      4,             // row with field_name keys (row 4 in your sheet)
  dataStartRow:   5,             // first character row
  // Column keys (matched against row 4 values)
  charNameKey:    'char',
  featsKey:       'feats',
  skillsKey:      'skills',
  classKeys:      ['class_1', 'class_2', 'class_3'],
  levelKey:       'level',
  classFeatures:     'class_features',	
  specialAbilities:  'special_abilities',
};
