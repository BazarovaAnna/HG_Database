function showGameRewardCorrection() {
  const ui = SpreadsheetApp.getUi();
  const range = SpreadsheetApp.getActiveRange();
  if (!range || range.getSheet().getName() !== 'Games' ||
      range.getRow() < 2 || range.getNumRows() !== 1) {
    ui.alert('Выберите одну строку партии на листе Games.');
    return;
  }
  try {
    const page = HtmlService.createTemplateFromFile('Game Reward Sidebar');
    const game = readGameCorrection_(range.getRow());
    page.initial = { row: game.row, revision: game.revision, fields: game.fields,
      processed: game.entries.length > 0 };
    ui.showSidebar(page.evaluate().setTitle('Исправить партию'));
  } catch (error) {
    ui.alert(error.message);
  }
}

function previewGameCorrection(row, revision, input) {
  const game = readGameCorrection_(row);
  return planGameCorrection_(game, revision, input).preview;
}

function applyGameCorrection(row, revision, input) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const game = readGameCorrection_(row);
    const plan = planGameCorrection_(game, revision, input);
    const date = Utilities.formatDate(
      new Date(), game.ss.getSpreadsheetTimeZone(), 'dd.MM.yyyy HH:mm:ss'
    );
    const requests = plan.changes.map(change => {
      const entry = date + ' — исправлено мастером: ' +
        (change.field.value || 'пусто') + ' → ' + (change.display || 'пусто') + '.';
      const note = change.field.note ? change.field.note + '\n' + entry : entry;
      const cell = {
        userEnteredValue: gameCellValue_(change.serial ?? change.value),
        note,
      };
      const fields = ['userEnteredValue', 'note'];
      if (change.serial !== undefined) {
        cell.userEnteredFormat = { numberFormat: { type: 'DATE', pattern: 'dd.MM.yyyy' } };
        fields.push('userEnteredFormat.numberFormat');
      }
      return gameCellsUpdate_(game.sheet, row, change.field.column, [[cell]], fields.join(','));
    });

    if (plan.ledgerChanged) {
      requests.push(...correctionTransactionUpdates_(game, plan.transactions, 'game'));
    }

    Sheets.Spreadsheets.batchUpdate({ requests }, game.ss.getId());
    if (plan.ledgerChanged) {
      try {
        buildInventorySheet(game.ss, game.txns);
        SpreadsheetApp.flush();
      } catch (error) {
        console.error(error);
        return 'Правка сохранена. Для обновления Inventory запустите PF1e Tools → Rebuild Inventory Sheet.';
      }
    }
    return 'Правка сохранена. История добавлена в примечания к изменённым полям.';
  } finally {
    lock.releaseLock();
  }
}

function readGameCorrection_(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Games');
  if (!sheet || !Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) {
    throw new Error('Строка партии недоступна.');
  }
  const count = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, count).getValues()[0];
  const range = sheet.getRange(row, 1, 1, count);
  const values = range.getValues()[0];
  const notes = range.getNotes()[0];
  const formulas = range.getFormulas()[0];
  const columns = loadColumnMappings(ss).games;
  const definitions = [
    ['played', 'Дата партии', 'date', headers.indexOf('When played') + 1],
    ['gameName', 'Название партии', 'text', columns.gameName],
    ['characters', 'Участники', 'list', columns.characters],
    ['exp', 'Опыт на участника', 'number', columns.exp],
    ['goldReward', 'Золото на участника', 'number', columns.goldReward],
    ['goldDelta', 'Другие изменения золота', 'multiline', columns.goldDelta],
    ['loot', 'Предметы', 'multiline', columns.loot],
    ['master', 'Мастер', 'text', headers.indexOf('Master') + 1],
  ];
  const indices = [columns.timestamp].concat(definitions.map(field => field[3]));
  if (new Set(indices).size !== indices.length ||
      !indices.every(column => Number.isInteger(column) && column > 0 && column <= count)) {
    throw new Error('Проверьте сопоставление колонок Games на листе Column matchups.');
  }
  const timestamp = String(getField(values, columns, 'timestamp'));
  if (!timestamp) throw new Error('В строке нет ответа формы.');
  const timestamps = sheet.getRange(2, columns.timestamp, sheet.getLastRow() - 1, 1).getValues();
  if (timestamps.filter(value => String(value[0]) === timestamp).length !== 1) {
    throw new Error('Отметка времени совпадает у нескольких партий.');
  }

  const fields = definitions.map(([key, label, type, column]) => {
    const raw = values[column - 1];
    const value = type === 'date' && raw instanceof Date
      ? Utilities.formatDate(raw, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd')
      : String(raw);
    return { key, label, type, column, value, note: notes[column - 1], formula: formulas[column - 1] };
  });
  const txns = ss.getSheetByName('Transactions');
  const ledger = txns && txns.getLastRow() > 1
    ? txns.getRange(2, 1, txns.getLastRow() - 1, TXN_HEADERS.length).getValues() : [];
  const entries = ledger.map((values, index) => ({ row: index + 2, values }))
    .filter(entry => String(entry.values[0]) === timestamp && entry.values[1] === 'game');
  const catalog = loadItemCatalog(ss);
  const revision = Utilities.base64EncodeWebSafe(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify([sheet.getSheetId(), txns?.getSheetId(), row, values, fields, ledger, catalog])
  ));
  return { ss, sheet, txns, row, columns, values, timestamp, fields, entries, ledger, catalog, revision };
}

function planGameCorrection_(game, revision, input) {
  if (game.revision !== revision) {
    throw new Error('Данные изменились. Откройте панель заново и проверьте правку.');
  }
  const values = game.values.slice();
  const changes = game.fields.filter(field => Object.hasOwn(input, field.key))
    .map(field => parseGameCorrectionField_(field, input[field.key]))
    .filter(change => change.display !== change.field.value);
  if (!changes.length) throw new Error('Поля не изменились.');
  changes.forEach(change => {
    if (change.field.formula) throw new Error('Поле содержит формулу: ' + change.field.label);
    values[change.field.column - 1] = change.value;
  });
  const changed = new Set(changes.map(change => change.field.key));
  const G = game.columns;
  const oldParticipants = gameParticipants_(getField(game.values, G, 'characters'));
  const participants = gameParticipants_(getField(values, G, 'characters'));
  const relevant = ['gameName', 'characters', 'goldReward', 'goldDelta', 'loot']
    .some(key => changed.has(key));
  const oldTransactions = relevant
    ? parseGameForm(game.values, G, game.catalog).filter(txn => txn.quantity !== 0) : [];
  const transactions = relevant
    ? parseGameForm(values, G, game.catalog).filter(txn => txn.quantity !== 0) : [];
  if (changed.has('characters') || changed.has('goldDelta') || changed.has('loot')) {
    validateGameCorrection_(game, values, participants, changed, transactions);
  }
  const ledgerChanged = relevant && game.entries.length > 0 &&
    JSON.stringify(oldTransactions.map(gameTransactionValues_)) !==
      JSON.stringify(transactions.map(gameTransactionValues_));
  if (ledgerChanged) validateGameLedger_(game, oldTransactions, transactions);

  const oldExperience = Number(getField(game.values, G, 'exp')) || 0;
  const experience = Number(getField(values, G, 'exp')) || 0;
  const effects = gameTransactionEffects_(oldTransactions, transactions, game.catalog);
  if (changed.has('characters') || changed.has('exp')) {
    for (const name of new Set([...oldParticipants, ...participants])) {
      const delta = (participants.includes(name) ? experience : 0) -
        (oldParticipants.includes(name) ? oldExperience : 0);
      if (delta) effects.push({ name, item: 'опыт', delta });
    }
  }
  return { changes, transactions, ledgerChanged, preview: {
    changes: changes.map(change => ({ label: change.field.label,
      from: change.field.value, to: change.display })),
    effects, pending: relevant && !game.entries.length,
    namedRewards: changed.has('characters') && ['goldDelta', 'loot'].some(key =>
      !changed.has(key) && String(getField(values, G, key)).trim()),
  } };
}

function parseGameCorrectionField_(field, input) {
  if (typeof input !== 'string') throw new Error('Проверьте поле «' + field.label + '».');
  const display = input.trim();
  if (display === field.value) return { field, display, value: display };
  if (field.type === 'number') {
    const value = display === '' ? '' : Number(display.replace(',', '.'));
    if (display && (!/^\d+([.,]\d{1,2})?$/.test(display) ||
        !Number.isSafeInteger(Math.round(value * 100)))) {
      throw new Error('Введите неотрицательное число в поле «' + field.label + '».');
    }
    return { field, display: String(value), value };
  }
  if (field.type === 'date' && display) {
    const date = new Date(display + 'T00:00:00Z');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(display) || !Number.isFinite(date.getTime()) ||
        date.toISOString().slice(0, 10) !== display) {
      throw new Error('Введите корректную дату партии.');
    }
    return { field, display, value: display, serial: (date.getTime() - Date.UTC(1899, 11, 30)) / 86400000 };
  }
  return { field, display, value: display };
}

function gameParticipants_(value) {
  return String(value).split(',').map(name => name.trim()).filter(Boolean);
}

function validateGameCorrection_(game, values, participants, changed, transactions) {
  const known = new Set(loadCharacterNames(game.ss));
  if (changed.has('characters') && (!participants.length ||
      new Set(participants).size !== participants.length ||
      participants.some(name => !known.has(name)))) {
    throw new Error('Укажите имена персонажей из Database через запятую, каждое один раз.');
  }
  for (const key of ['goldDelta', 'loot']) {
    if (!changed.has(key)) continue;
    const text = String(getField(values, game.columns, key));
    for (const block of text.split(';').map(value => value.trim()).filter(Boolean)) {
      const colon = block.indexOf(':');
      const character = block.slice(0, colon).trim();
      const payload = block.slice(colon + 1).trim();
      const valid = key === 'goldDelta'
        ? /^[+-]?\d+(\.\d{1,2})?$/.test(payload) && Number.isSafeInteger(Math.round(Number(payload) * 100))
        : payload.split(',').every(item => parseItemFull(item));
      if (colon < 1 || !known.has(character) || !valid) {
        throw new Error('Проверьте имя персонажа и значение в записи «' + block + '».');
      }
    }
  }
  if (transactions.some(txn => !Number.isFinite(txn.quantity) ||
      !Number.isFinite(txn.cost) || (txn.weight !== '' && !Number.isFinite(txn.weight)))) {
    throw new Error('Проверьте количество, цену и вес предметов.');
  }
}

function validateGameLedger_(game, previous, next, checkGoldTransfers) {
  const start = game.entries[0].row;
  const count = game.entries.length;
  const formulas = game.txns.getRange(start, 1, count, TXN_HEADERS.length).getFormulas();
  if (game.entries.some((entry, index) => entry.row !== start + index) ||
      formulas.some(row => row.some(Boolean))) {
    throw new Error('Записи партии расположены в разных местах или содержат формулы.');
  }
  const normalize = rows => rows.filter(row => row[2] && row[3] && Number(row[4]) !== 0)
    .map(row => JSON.stringify(row.slice(2))).sort();
  if (JSON.stringify(normalize(game.entries.map(entry => entry.values))) !==
      JSON.stringify(normalize(previous.map(gameTransactionValues_)))) {
    throw new Error('Записи в Transactions расходятся с исходным ответом. Проверьте журнал.');
  }

  const itemEffects = gameTransactionEffects_(previous, next, game.catalog)
    .filter(effect => checkGoldTransfers || effect.item.toLowerCase() !== 'gold');
  const tracked = new Map(itemEffects.map(effect =>
    [inventoryItemKey(effect.name, effect.item), { effect, before: 0, after: 0 }]));
  const add = (row, column) => {
    if (!row[2] || !row[3]) return;
    const balance = tracked.get(inventoryItemKey(row[2], canonicalItemName(row[3], game.catalog)));
    if (balance) balance[column] += Number(row[4]) || 0;
  };
  const check = row => {
    for (const { before, after, effect } of tracked.values()) {
      if (effect.item.toLowerCase() === 'gold') {
        const entry = game.ledger[row - 2];
        if (row < start + count || entry[1] !== 'transfer' || entry[2] !== effect.name ||
            String(entry[3]).toLowerCase() !== 'gold' || Number(entry[4]) >= 0) continue;
      }
      if (after < -0.000001 && after < before - 0.000001) {
        throw new Error('После исправления не хватает предмета «' + effect.item +
          '» у ' + effect.name + '. Проверьте Transactions, строка ' + row + '.');
      }
    }
  };
  for (let i = 0; i < game.ledger.length; i++) {
    const row = i + 2;
    add(game.ledger[i], 'before');
    if (row === start) next.forEach(txn => add(gameTransactionValues_(txn), 'after'));
    if (row < start || row >= start + count) add(game.ledger[i], 'after');
    if (row >= start + count - 1) check(row);
  }
}

function correctionTransactionUpdates_(correction, transactions, markerSource) {
  const start = correction.entries[0].row;
  const existing = correction.entries.length;
  const replacement = transactions.map(gameTransactionValues_);
  const requests = [];
  const extra = replacement.length - existing;
  if (extra > 0) {
    const startIndex = start + existing - 1;
    const endIndex = startIndex + extra;
    const sheetId = correction.txns.getSheetId();
    const gridRows = correction.txns.getMaxRows();
    if (endIndex > gridRows) {
      requests.push({ appendDimension: { sheetId, dimension: 'ROWS', length: endIndex - gridRows } });
    }
    requests.push({ insertDimension: {
      range: { sheetId, dimension: 'ROWS', startIndex, endIndex }, inheritFromBefore: true,
    } });
  }
  while (replacement.length < existing) {
    replacement.push([correction.timestamp, markerSource, '', '', '', '', '', '']);
  }
  requests.push(gameCellsUpdate_(correction.txns, start, 1,
    replacement.map(values => values.map(value => ({ userEnteredValue: gameCellValue_(value) }))),
    'userEnteredValue'));
  return requests;
}

function gameTransactionEffects_(previous, next, catalog) {
  const effects = new Map();
  const accumulate = (transactions, sign) => transactions.forEach(txn => {
    const name = txn.character;
    const item = canonicalItemName(txn.itemName, catalog);
    const key = inventoryItemKey(name, item);
    const effect = effects.get(key) || { name, item, delta: 0 };
    effect.delta += sign * txn.quantity;
    effects.set(key, effect);
  });
  accumulate(previous, -1);
  accumulate(next, 1);
  return [...effects.values()].filter(effect => Math.abs(effect.delta) > 0.000001);
}

function gameTransactionValues_(txn) {
  return [txn.timestamp, txn.source, txn.character, txn.itemName,
    txn.quantity, txn.cost, txn.weight, txn.notes];
}

function gameCellValue_(value) {
  if (value === '') return undefined;
  return typeof value === 'number' ? { numberValue: value } : { stringValue: value };
}

function gameCellsUpdate_(sheet, row, column, cells, fields) {
  return { updateCells: {
    start: { sheetId: sheet.getSheetId(), rowIndex: row - 1, columnIndex: column - 1 },
    rows: cells.map(values => ({ values })),
    fields,
  } };
}
