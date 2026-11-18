function showInventoryCorrection() {
  const ui = SpreadsheetApp.getUi();
  const range = SpreadsheetApp.getActiveRange();
  if (!range || range.getSheet().getName() !== 'Inventory(raw)' ||
      range.getRow() < 2 || range.getNumRows() !== 1) {
    ui.alert('Выберите строку заявки на листе Inventory(raw).');
    return;
  }
  try {
    const request = readInventoryCorrection_(range.getRow());
    const page = HtmlService.createTemplateFromFile('Game Reward Sidebar');
    page.initial = { kind: 'inventory', row: request.row, revision: request.revision,
      fields: request.fields, context: 'Inventory(raw) · строка ' + request.row + ' · ' +
        (request.entries.length ? 'обработана' : 'ещё не обработана') };
    ui.showSidebar(page.evaluate().setTitle('Исправить заявку на инвентарь'));
  } catch (error) {
    ui.alert(error.message);
  }
}

function previewInventoryCorrection(row, revision, input) {
  return planInventoryCorrection_(readInventoryCorrection_(row), revision, input).preview;
}

function applyInventoryCorrection(row, revision, input) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const request = readInventoryCorrection_(row);
    const plan = planInventoryCorrection_(request, revision, input);
    const date = Utilities.formatDate(new Date(), request.ss.getSpreadsheetTimeZone(), 'dd.MM.yyyy HH:mm:ss');
    const updates = plan.changes.map(change => {
      const entry = date + ' — исправлено мастером: ' + (change.field.value || 'пусто') +
        ' → ' + (change.display || 'пусто') + '.';
      return gameCellsUpdate_(request.sheet, row, change.field.column, [[{
        userEnteredValue: gameCellValue_(change.value),
        note: change.field.note ? change.field.note + '\n' + entry : entry,
      }]], 'userEnteredValue,note');
    });
    if (plan.ledgerChanged) {
      updates.push(...correctionTransactionUpdates_(request, plan.result.txns, 'inventory'));
    }
    Sheets.Spreadsheets.batchUpdate({ requests: updates }, request.ss.getId());
    if (plan.ledgerChanged) {
      try {
        buildInventorySheet(request.ss, request.txns);
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

function readInventoryCorrection_(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Inventory(raw)');
  if (!sheet || !Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) {
    throw new Error('Строка заявки недоступна.');
  }
  const count = sheet.getLastColumn();
  const columns = loadColumnMappings(ss).inventory;
  const definitions = [
    ['charName', 'Персонаж', 'text'], ['whatHappened', 'Что произошло', 'text'],
    ['whoWith', 'С кем', 'text'], ['worldItems', 'Предметы и золото', 'multiline'],
    ['participants', 'Участники передачи', 'text'], ['transfers', 'Передачи', 'multiline'],
    ['comment', 'Комментарий из формы', 'multiline'],
  ];
  const indices = [columns.timestamp, columns.approved].concat(definitions.map(field => columns[field[0]]));
  if (new Set(indices).size !== indices.length ||
      !indices.every(column => Number.isInteger(column) && column > 0 && column <= count)) {
    throw new Error('Проверьте сопоставление колонок Inventory(raw) на листе Column matchups.');
  }
  const range = sheet.getRange(row, 1, 1, count);
  const values = range.getValues()[0];
  const notes = range.getNotes()[0];
  const formulas = range.getFormulas()[0];
  const timestamp = String(getField(values, columns, 'timestamp'));
  if (!timestamp) throw new Error('В строке нет ответа формы.');
  const timestamps = sheet.getRange(2, columns.timestamp, sheet.getLastRow() - 1, 1).getValues();
  if (timestamps.filter(value => String(value[0]) === timestamp).length !== 1) {
    throw new Error('Отметка времени совпадает у нескольких заявок.');
  }
  const fields = definitions.map(([key, label, type]) => {
    const column = columns[key];
    return { key, label, type, column, value: String(values[column - 1]),
      note: notes[column - 1], formula: formulas[column - 1] };
  });
  const txns = ss.getSheetByName('Transactions');
  const ledger = txns && txns.getLastRow() > 1
    ? txns.getRange(2, 1, txns.getLastRow() - 1, TXN_HEADERS.length).getValues() : [];
  const entries = ledger.map((values, index) => ({ values, row: index + 2 }))
    .filter(entry => String(entry.values[0]) === timestamp && REQUEST_SOURCES.includes(entry.values[1]));
  const catalog = loadItemCatalog(ss);
  const revision = Utilities.base64EncodeWebSafe(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify([sheet.getSheetId(), txns?.getSheetId(), row, values, fields, ledger, catalog])
  ));
  return { ss, sheet, txns, row, columns, values, timestamp, fields, entries, ledger, catalog, revision };
}

function planInventoryCorrection_(request, revision, input) {
  if (request.revision !== revision) {
    throw new Error('Данные изменились. Откройте панель заново и проверьте правку.');
  }
  const changes = request.fields.filter(field => Object.hasOwn(input, field.key))
    .map(field => parseGameCorrectionField_(field, input[field.key]))
    .filter(change => change.display !== change.field.value);
  if (!changes.length) throw new Error('Поля не изменились.');
  const values = request.values.slice();
  changes.forEach(change => {
    if (change.field.formula) throw new Error('Поле содержит формулу: ' + change.field.label);
    values[change.field.column - 1] = change.value;
  });
  const relevant = changes.some(change => change.field.key !== 'comment');
  const preview = {
    changes: changes.map(change => ({ label: change.field.label,
      from: change.field.value, to: change.display })),
    effects: [], warnings: [],
  };
  if (!relevant) return { changes, ledgerChanged: false, preview };

  const I = request.columns;
  const characters = loadCharacterNames(request.ss);
  const character = String(getField(values, I, 'charName')).trim();
  if (!characters.includes(character)) throw new Error('Укажите имя персонажа как на листе Database.');
  const branch = String(getField(values, I, 'whoWith')).trim();
  const world = String(getField(values, I, 'worldItems')).trim();
  const transfers = String(getField(values, I, 'transfers')).trim();
  if ((branch === 'The world' && transfers) || (branch === 'Other characters' && world) ||
      (world && transfers)) {
    throw new Error('Оставьте заполненным поле, которое соответствует выбранному варианту «С кем».');
  }
  if (changes.some(change => change.field.key === 'whoWith') &&
      !['The world', 'Other characters'].includes(branch)) {
    throw new Error('В поле «С кем» укажите The world или Other characters.');
  }

  const processed = request.entries.length > 0;
  const start = processed ? request.entries[0].row : request.ledger.length + 2;
  const inventory = {};
  request.ledger.slice(0, start - 2).forEach(row => {
    if (!row[2] || !row[3]) return;
    updateInventoryMap(inventory, { character: row[2],
      itemName: canonicalItemName(row[3], request.catalog), quantity: Number(row[4]) || 0,
      cost: row[5], weight: row[6] });
  });
  const revisedInventory = copyInventoryMap(inventory);
  const result = parseInventoryRequest(values, I, revisedInventory, character, request.catalog, characters);
  if (result.errors.length) throw new Error(result.errors.join('\n'));
  if (result.txns.some(txn => !Number.isFinite(txn.quantity))) {
    throw new Error('Проверьте количество золота и предметов.');
  }

  let previous = [];
  let ledgerChanged = false;
  if (processed) {
    const characterBefore = String(getField(request.values, I, 'charName')).trim();
    const original = parseInventoryRequest(request.values, I, inventory, characterBefore,
      request.catalog, characters);
    if (original.errors.length) {
      throw new Error('Не удалось сверить исходную заявку: ' + original.errors.join('\n'));
    }
    previous = original.txns;
    ledgerChanged = JSON.stringify(previous.map(gameTransactionValues_)) !==
      JSON.stringify(result.txns.map(gameTransactionValues_));
    if (ledgerChanged) {
      validateGameLedger_(request, previous, result.txns, true);
      validateInventoryCorrectionValues_(request, start, inventory, revisedInventory);
    }
  }
  preview.effects = gameTransactionEffects_(previous, result.txns, request.catalog);
  preview.warnings = result.warnings;
  if (!processed) {
    preview.publication = isApproved(getField(values, I, 'approved'))
      ? 'Золото и предметы обновятся при обработке заявки.'
      : 'Транзакции появятся после утверждения и обработки заявки.';
  }
  return { changes, result, ledgerChanged, preview };
}

function validateInventoryCorrectionValues_(request, start, original, revised) {
  const last = start + request.entries.length - 1;
  for (let index = last - 1; index < request.ledger.length; index++) {
    const entry = request.ledger[index];
    if (!entry[2] || !entry[3] || Number(entry[4]) >= 0 || String(entry[3]).toLowerCase() === 'gold') continue;
    const item = canonicalItemName(entry[3], request.catalog);
    const key = inventoryItemKey(entry[2], item);
    if (original[key] && revised[key] &&
        (original[key].cost !== revised[key].cost || original[key].weight !== revised[key].weight)) {
      throw new Error('Правка меняет цену или вес предмета «' + item +
        '», который позже выбыл. Проверьте Transactions, строка ' + (index + 2) + '.');
    }
  }
}
