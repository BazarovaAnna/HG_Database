const CHARACTER_ARCHIVED_STATUS = 'Архив';
const NO_ACTIVE_CHARACTER_CHOICE = 'Нет активных персонажей';
const CHARACTER_MANUAL_COLUMNS = [
  'str_racial_var', 'dex_racial_var', 'con_racial_var', 'int_racial_var', 'wis_racial_var', 'cha_racial_var',
  'speed_misc', 'init_misc', 'ac_misc', 'fort_misc', 'ref_misc', 'will_misc',
];

function showCharacterRemoval() {
  const ui = SpreadsheetApp.getUi();
  const range = SpreadsheetApp.getActiveRange();
  if (!range || range.getSheet().getName() !== 'Registration' ||
      range.getRow() < 2 || range.getNumRows() !== 1) {
    ui.alert('Выберите строку персонажа на листе Registration.');
    return;
  }
  try {
    const request = readCharacterRemoval_(range.getRow());
    const page = HtmlService.createTemplateFromFile('Character Removal Sidebar');
    page.initial = { row: request.row, name: request.name, archived: request.archived };
    ui.showSidebar(page.evaluate().setTitle('Remove character'));
  } catch (error) {
    ui.alert(error.message);
  }
}

function previewCharacterRemoval(row, mode, expectedName) {
  const request = readCharacterRemoval_(row);
  if (request.name !== expectedName) {
    throw new Error('В выбранной строке теперь другой персонаж. Откройте панель заново.');
  }
  const plan = planCharacterRemoval_(request, mode);
  return { revision: request.revision, summary: plan.summary };
}

function applyCharacterRemoval(row, mode, revision, confirmation) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const request = readCharacterRemoval_(row);
    if (request.revision !== revision) {
      throw new Error('Данные изменились. Нажмите «Проверить» ещё раз.');
    }
    if (mode === 'delete' && confirmation !== request.name) {
      throw new Error('Введите короткое имя персонажа точно как в Registration.');
    }
    const plan = planCharacterRemoval_(request, mode);
    Sheets.Spreadsheets.batchUpdate({ requests: plan.updates }, request.ss.getId());

    const pending = [];
    const refresh = (command, operation) => {
      try {
        SpreadsheetApp.flush();
        operation();
      } catch (error) {
        console.error(error);
        pending.push(command);
      }
    };
    refresh('Refresh ALL normalized tables', () => { parseFeats(); parseSkills(); parseSpells(); });
    if (mode === 'delete') {
      refresh('Rebuild Inventory Sheet', () => buildInventorySheet(request.ss, request.sheets.Transactions.sheet));
    }
    refresh('Sync form choices', () => {
      const report = syncFormChoices(true);
      if (report.errors.length || report.missing.length || report.wrongType.length) {
        throw new Error(JSON.stringify(report));
      }
    });
    refresh('Find data Discrepancies', () => findDiscrepancies(true));
    const result = mode === 'archive' ? 'Персонаж отправлен в архив.' : 'Записи персонажа удалены из таблицы.';
    return pending.length ? result + ' Повторите через PF1e Tools: ' + pending.join(', ') + '.' : result;
  } finally {
    lock.releaseLock();
  }
}

function readCharacterRemoval_(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const columns = loadColumnMappings(ss);
  const R = columns.registration;
  const required = ['Registration', 'Database', 'LvlUP', 'LevelUp Log', 'Transactions',
    'Games', 'Inventory(raw)', 'Bulletin Board', 'Inspector'];
  const sheets = Object.fromEntries(required.map(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error('Не найден лист «' + name + '».');
    const values = name === 'Inspector' ? sheet.getRange('B1').getValues() : sheet.getDataRange().getValues();
    return [name, { sheet, values }];
  }));
  const registrations = sheets.Registration.values;
  if (!Number.isInteger(row) || row < 2 || row > registrations.length ||
      registrations[0][R.approved - 1] !== 'Статус' || !R.charName) {
    throw new Error('Проверьте выбранную строку и сопоставление колонок Registration.');
  }
  const values = registrations[row - 1];
  const name = String(getField(values, R, 'charName')).trim();
  const status = String(getField(values, R, 'approved'));
  if (!name || registrations.slice(1).filter(record =>
    registrationNameMatches_(getField(record, R, 'charName'), name)).length !== 1) {
    throw new Error('В Registration должна быть ровно одна строка с этим коротким именем.');
  }
  if (registrations.slice(1).filter(record => nameKey(getField(record, R, 'charName')) === nameKey(name)).length !== 1) {
    throw new Error('У другого персонажа похожее короткое имя. Сначала уточните имена.');
  }
  const activeNames = registrations.slice(1).filter(record => getField(record, R, 'approved') === 'Утверждено')
    .map(record => String(getField(record, R, 'charName')).trim());
  const database = sheets.Database;
  const headers = database.values[3] || [];
  const charColumn = headers.indexOf('char');
  const actual = database.values.slice(4).map(record => String(record[charColumn] ?? '').trim());
  const activeIndex = activeNames.indexOf(name);
  if (charColumn < 0 || activeNames.some((name, index) => !name || name !== actual[index]) ||
      actual.slice(activeNames.length).some(Boolean)) {
    throw new Error('Порядок персонажей в Database расходится с Registration. Проверьте формулы и статусы.');
  }
  const manual = CHARACTER_MANUAL_COLUMNS.map(key => headers.indexOf(key));
  if (!manual.every(index => index >= 0)) {
    throw new Error('В Database изменились заголовки ручных колонок.');
  }
  let databaseCells = [];
  if (activeIndex >= 0) {
    const range = database.sheet.getRange(5, 1, activeNames.length, headers.length);
    const result = Sheets.Spreadsheets.get(ss.getId(), {
      ranges: ['Database!' + range.getA1Notation()],
      fields: 'sheets(data(rowData(values(userEnteredValue,note))))',
    });
    databaseCells = result.sheets?.[0]?.data?.[0]?.rowData || [];
  }
  const statusCell = sheets.Registration.sheet.getRange(row, R.approved);
  const rule = statusCell.getDataValidation();
  const validStatuses = rule && rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST
    ? rule.getCriteriaValues()[0].map(String) : null;
  const statusNote = statusCell.getNote();
  const revision = Utilities.base64EncodeWebSafe(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, JSON.stringify([row, columns, statusNote, validStatuses,
      required.map(key => [sheets[key].sheet.getSheetId(), sheets[key].values]), databaseCells])
  ));
  return { ss, row, name, status, archived: status === CHARACTER_ARCHIVED_STATUS, columns,
    sheets, activeIndex, activeNames, manual, databaseCells, validStatuses, statusNote, revision };
}

function planCharacterRemoval_(request, mode) {
  if (mode !== 'archive' && mode !== 'delete') throw new Error('Выберите действие.');
  if (mode === 'archive' && request.archived) throw new Error('Этот персонаж уже в архиве.');
  const updates = [];
  const summary = [];
  const R = request.columns.registration;
  const registration = request.sheets.Registration;
  const manualBackup = moveCharacterDatabaseFields_(request, updates);
  if (request.activeIndex >= 0) {
    const after = request.activeNames.length - request.activeIndex - 1;
    summary.push('Персонаж исчезнет из Database и списка выбора. Ручные поля ' +
      after + ' следующих карточек будут перемещены вместе с ними.');
  }
  if (mode === 'archive') {
    if (!request.validStatuses) throw new Error('У статуса регистрации нужен список допустимых значений.');
    const date = Utilities.formatDate(new Date(), request.ss.getSpreadsheetTimeZone(), 'dd.MM.yyyy HH:mm:ss');
    const heading = date + ' — персонаж архивирован мастером. Прежний статус: ' + (request.status || 'пусто') + '.';
    const note = [request.statusNote, heading,
      Object.keys(manualBackup).length ? 'Ручные поля Database: ' + JSON.stringify(manualBackup) : '']
      .filter(Boolean).join('\n');
    const range = characterRemovalRange_(registration.sheet, request.row, R.approved, R.approved);
    const statuses = [...new Set([...request.validStatuses, CHARACTER_ARCHIVED_STATUS])];
    updates.push({ setDataValidation: { range, rule: {
      condition: { type: 'ONE_OF_LIST', values: statuses.map(userEnteredValue => ({ userEnteredValue })) },
      strict: true, showCustomUi: true,
    } } });
    updates.push(gameCellsUpdate_(registration.sheet, request.row, R.approved, [[{
      userEnteredValue: gameCellValue_(CHARACTER_ARCHIVED_STATUS), note,
    }]], 'userEnteredValue,note'));
    summary.push('Регистрация получит статус «Архив». Ответы, история партий и транзакции сохранятся.');
    if (Object.keys(manualBackup).length) summary.push('Ручные поля персонажа будут скопированы в примечание к статусу.');
  } else {
    clearCharacterRow_(registration, request.row, updates);
    summary.push('Регистрация и личные записи персонажа будут удалены из книги.');
    summarizeCharacterPurge_(request, updates, summary);
    summary.push('Опыт, золото и предметы других участников сохранятся. Ответы в Google Forms и история версий Google Sheets сохраняются.');
  }
  const inspector = request.sheets.Inspector;
  if (registrationNameMatches_(inspector.values[0][0], request.name)) {
    writeCharacterRemovalCell_(inspector, 1, 2, '', updates);
  }
  return { updates, summary };
}

function moveCharacterDatabaseFields_(request, updates) {
  const backup = {};
  if (request.activeIndex < 0) return backup;
  const database = request.sheets.Database;
  const manual = new Set(request.manual);
  const cells = (row, column) => request.databaseCells[row]?.values?.[column] || {};
  for (let row = request.activeIndex; row < request.activeNames.length; row++) {
    const entries = request.databaseCells[row]?.values || [];
    if (entries.some((cell, index) => (cell.userEnteredValue || cell.note) && !manual.has(index))) {
      throw new Error('На Database, строка ' + (row + 5) + ', есть ручные данные вне известных колонок.');
    }
  }
  for (const index of request.manual) {
    const removed = cells(request.activeIndex, index);
    if (removed.userEnteredValue || removed.note) backup[database.values[3][index]] = removed;
    const values = [];
    for (let row = request.activeIndex; row < request.activeNames.length; row++) {
      const current = cells(row, index);
      if (current.userEnteredValue?.formulaValue) {
        throw new Error('Ручное поле Database содержит формулу, строка ' + (row + 5) + '.');
      }
      values.push([cells(row + 1, index)]);
    }
    updates.push(gameCellsUpdate_(database.sheet, request.activeIndex + 5, index + 1,
      values, 'userEnteredValue,note'));
  }
  return backup;
}

function summarizeCharacterPurge_(request, updates, summary) {
  const sources = request.sheets;
  const { lvlup, games, inventory } = request.columns;
  const boardAuthor = sources['Bulletin Board'].values[0].indexOf('Author') + 1;
  const required = [lvlup.charName, games.timestamp, games.characters, games.goldDelta, games.loot,
    inventory.charName, inventory.participants, inventory.transfers, inventory.worldItems,
    inventory.timestamp, boardAuthor];
  if (required.some(column => !Number.isInteger(column) || column < 1)) {
    throw new Error('Проверьте сопоставление колонок связанных форм.');
  }
  const name = request.name;
  const matches = value => registrationNameMatches_(value, name);
  const counts = [];
  const ownRows = (sheetName, column) => {
    const source = sources[sheetName];
    let count = 0;
    source.values.slice(1).forEach((values, index) => {
      if (matches(values[column - 1])) {
        clearCharacterRow_(source, index + 2, updates);
        count++;
      }
    });
    if (count) counts.push(sheetName + ' — ' + count);
  };
  ownRows('LvlUP', lvlup.charName);
  ownRows('LevelUp Log', LOG.character);
  ownRows('Bulletin Board', boardAuthor);
  ownRows('Transactions', TXN.character);

  const transactionRows = sources.Transactions.values.slice(1);
  const counterparty = value => String(value).replace(/^(To |From )([^:]+)(?=:|$)/,
    (original, prefix, character) => matches(character) ? prefix + 'удалённый персонаж' : original);
  transactionRows.forEach((values, index) => {
    if (matches(values[TXN.character - 1]) || values[TXN.source - 1] !== 'transfer') return;
    const note = values[TXN.notes - 1];
    const value = counterparty(note);
    if (value !== note) writeCharacterRemovalCell_(sources.Transactions, index + 2, TXN.notes, value, updates);
  });

  const filterEntries = (value, separator, joiner, keep) => {
    const entries = String(value).split(separator);
    const remaining = entries.filter(keep);
    return remaining.length === entries.length ? value : remaining.join(joiner).trim();
  };
  const removeFromList = value => filterEntries(value, ',', ',', entry => !matches(entry));
  const removeNamedReward = value => filterEntries(value, ';', ';', entry => {
    const colon = entry.indexOf(':');
    return colon < 0 || !matches(entry.slice(0, colon));
  });
  let gamesChanged = 0;
  sources.Games.values.slice(1).forEach((values, index) => {
    const row = index + 2;
    const changes = [
      [games.characters, removeFromList], [games.goldDelta, removeNamedReward], [games.loot, removeNamedReward],
    ].map(([column, transform]) => [column, transform(values[column - 1] ?? '')]);
    if (!changes.some(([column, value]) => value !== values[column - 1])) return;
    const [participants, extraGold, loot] = changes.map(change => change[1]);
    const timestamp = String(getField(values, games, 'timestamp'));
    if (!participants && (extraGold || loot) && !transactionRows.some(record =>
      String(record[0]) === timestamp && record[1] === 'game')) {
      throw new Error('Сначала обработайте партию в Games, строка ' + row + ': награды других персонажей ещё не начислены.');
    }
    changes.forEach(([column, value]) => {
      if (value !== values[column - 1]) writeCharacterRemovalCell_(sources.Games, row, column, value, updates);
    });
    gamesChanged++;
  });

  let inventoryChanged = 0;
  const alias = value => nameKey(value) === nameKey(name);
  const removeTransfers = value => filterEntries(value, /\n|;/, '\n', line => {
    const match = line.trim().match(/^(.+?)\s*(?:->|→|=>|>)\s*(.+?)\s*:/);
    return !match || (!alias(match[1]) && !alias(match[2]));
  });
  sources['Inventory(raw)'].values.slice(1).forEach((values, index) => {
    const row = index + 2;
    const author = getField(values, inventory, 'charName');
    const participants = getField(values, inventory, 'participants');
    const transfers = getField(values, inventory, 'transfers');
    const remaining = removeTransfers(transfers);
    const remainingParticipants = removeFromList(participants);
    const source = sources['Inventory(raw)'];
    if (matches(author) && !remaining) {
      clearCharacterRow_(source, row, updates);
    } else {
      if (!matches(author) && remaining === transfers && remainingParticipants === participants) return;
      if (matches(author)) {
        writeCharacterRemovalCell_(source, row, inventory.charName, 'Удалённый персонаж', updates);
        writeCharacterRemovalCell_(source, row, inventory.worldItems, '', updates);
      }
      if (remaining !== transfers) writeCharacterRemovalCell_(source, row, inventory.transfers, remaining, updates);
      if (remainingParticipants !== participants) writeCharacterRemovalCell_(source, row, inventory.participants, remainingParticipants, updates);
    }
    inventoryChanged++;
  });
  if (counts.length) summary.push('Личные строки будут очищены: ' + counts.join(', ') + '.');
  if (gamesChanged || inventoryChanged) {
    summary.push('Связанные строки будут обновлены: Games — ' + gamesChanged + ', Inventory(raw) — ' + inventoryChanged + '.');
  }
}

function characterRemovalRange_(sheet, row, start, end) {
  return { sheetId: sheet.getSheetId(), startRowIndex: row - 1, endRowIndex: row,
    startColumnIndex: start - 1, endColumnIndex: end };
}

function clearCharacterRow_(source, row, updates) {
  const width = source.values[0].length;
  if (source.sheet.getRange(row, 1, 1, width).getFormulas()[0].some(Boolean)) {
    throw new Error('Связанная строка содержит формулу: ' + source.sheet.getName() + ' · ' + row + '.');
  }
  updates.push({ updateCells: {
    range: characterRemovalRange_(source.sheet, row, 1, width), fields: 'userEnteredValue,note',
  } });
}

function writeCharacterRemovalCell_(source, row, column, value, updates) {
  if (source.sheet.getRange(row, column).getFormula()) {
    throw new Error('Связанная ячейка содержит формулу: ' + source.sheet.getName() + ' · ' + row + '.');
  }
  updates.push(gameCellsUpdate_(source.sheet, row, column,
    [[{ userEnteredValue: gameCellValue_(value) }]], 'userEnteredValue'));
}

function archivedCharacterNames_(ss, columns) {
  const sheet = ss.getSheetByName('Registration');
  const R = columns || loadColumnMappings(ss).registration;
  if (!sheet || !R.approved || !R.charName) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter(row => getField(row, R, 'approved') === CHARACTER_ARCHIVED_STATUS)
    .map(row => String(getField(row, R, 'charName')).trim().toLowerCase())
    .filter(Boolean);
}
