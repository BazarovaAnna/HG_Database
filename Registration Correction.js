function showRegistrationCorrection() {
  const ui = SpreadsheetApp.getUi();
  const range = SpreadsheetApp.getActiveRange();
  if (!range || range.getSheet().getName() !== 'Registration' ||
      range.getRow() < 2 || range.getNumRows() !== 1) {
    ui.alert('Выберите строку на листе Registration.');
    return;
  }
  try {
    const registration = readRegistrationCorrection_(range.getRow());
    const page = HtmlService.createTemplateFromFile('Game Reward Sidebar');
    page.initial = { kind: 'registration', row: registration.row, revision: registration.revision,
      fields: registration.fields, context: 'Registration · строка ' + registration.row + ' · ' +
        (registration.approved ? 'утверждена' : 'ещё не утверждена') };
    ui.showSidebar(page.evaluate().setTitle('Исправить регистрацию'));
  } catch (error) {
    ui.alert(error.message);
  }
}

function previewRegistrationCorrection(row, revision, input) {
  return planRegistrationCorrection_(readRegistrationCorrection_(row), revision, input).preview;
}

function applyRegistrationCorrection(row, revision, input) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const request = readRegistrationCorrection_(row);
    const plan = planRegistrationCorrection_(request, revision, input);
    const now = new Date();
    const date = Utilities.formatDate(now, request.ss.getSpreadsheetTimeZone(), 'dd.MM.yyyy HH:mm:ss');
    const corrections = new Map(plan.changes.map(change => [change.field.column, change]));
    const registrationValues = new Map(plan.changes.map(change => [change.field.column, change.value]));
    plan.replay?.changes.forEach(change => registrationValues.set(change.column, change.after));
    // Replaying can absorb a changed initial value without altering the final card.
    plan.replay?.events.forEach(event => event.changes.forEach(change => {
      if (corrections.has(change.regCol) && !plan.replay.changes.some(result => result.column === change.regCol)) {
        registrationValues.set(change.regCol, request.values[change.regCol - 1]);
      }
    }));
    const updates = [...registrationValues].map(([column, value]) => {
      const change = corrections.get(column);
      const cell = { userEnteredValue: gameCellValue_(value) };
      if (change) {
        const line = date + ' — исправлено мастером: ' + (change.field.value || 'пусто') +
          ' → ' + (change.display || 'пусто') + '.';
        cell.note = change.field.note ? change.field.note + '\n' + line : line;
      }
      return gameCellsUpdate_(request.sheet, row, column, [[cell]], change ? 'userEnteredValue,note' : 'userEnteredValue');
    });
    updates.push(...plan.links.map(link => gameCellsUpdate_(link.sheet, link.row, link.column,
      [[{ userEnteredValue: gameCellValue_(link.value) }]], 'userEnteredValue')));

    if (plan.replay) {
      const sheetId = request.sheets['LevelUp Log'].sheet.getSheetId();
      plan.replay.oldLogs.slice().sort((a, b) => b - a).forEach(logRow => {
        updates.push({ deleteDimension: { range: {
          sheetId, dimension: 'ROWS', startIndex: logRow - 1, endIndex: logRow,
        } } });
      });
      const rows = plan.replay.events.flatMap(event => {
        const changes = event.changes.length ? event.changes : [{ field: '(no changes)', regCol: 0, old: '', new: '' }];
        return changes.map((change, index) => [
          event.timestamp, now.toISOString(), plan.name, event.level, change.field,
          change.regCol, change.old, change.new,
          index === 0 ? String(getField(event.values, request.columns.lvlup, 'gmNotes')).trim() : '',
        ]);
      });
      updates.push({ appendCells: { sheetId, fields: 'userEnteredValue',
        rows: rows.map(values => ({ values: values.map(value => ({
          userEnteredValue: gameCellValue_(value),
        })) })),
      } });
    }
    if (plan.newGold !== null) {
      const sheet = request.sheets.Transactions.sheet;
      const sheetId = sheet.getSheetId();
      updates.push({ insertDimension: { range: { sheetId, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
        inheritFromBefore: false } });
      const values = [now.toISOString(), 'registration', plan.name, 'gold', plan.newGold, 1, '', 'Starting gold'];
      updates.push(gameCellsUpdate_(sheet, 2, 1, [values.map(value => ({
        userEnteredValue: gameCellValue_(value),
      }))], 'userEnteredValue'));
    }

    Sheets.Spreadsheets.batchUpdate({ requests: updates }, request.ss.getId());
    const pending = [];
    const refresh = (command, operation) => {
      try { operation(); } catch (error) { console.error(error); pending.push(command); }
    };
    if (request.approved) refresh('Refresh ALL normalized tables', () => {
      SpreadsheetApp.flush();
      parseFeats();
      parseSkills();
      parseSpells();
    });
    if (plan.rename || plan.goldChanged) {
      refresh('Rebuild Inventory Sheet', () => buildInventorySheet(request.ss, request.sheets.Transactions.sheet));
    }
    if (plan.rename && request.approved) refresh('Sync form choices', () => {
      SpreadsheetApp.flush();
      const report = syncFormChoices(true);
      if (report.errors.length || report.missing.length || report.wrongType.length) {
        throw new Error(JSON.stringify(report));
      }
    });
    if (request.approved) refresh('Find data Discrepancies', () => {
      SpreadsheetApp.flush();
      findDiscrepancies(true);
    });
    return pending.length ? 'Правка сохранена. Повторите через PF1e Tools: ' + pending.join(', ') + '.'
      : 'Правка сохранена. История добавлена в примечания к изменённым полям.';
  } finally {
    lock.releaseLock();
  }
}

function readRegistrationCorrection_(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const columns = loadColumnMappings(ss);
  const R = columns.registration;
  const names = ['Registration', 'LvlUP', 'LevelUp Log', 'Transactions', 'Games', 'Inventory(raw)',
    'Bulletin Board', 'Database', 'Reference tables', 'Inventory', 'Feats', 'Spells', 'Inspector', 'Discrepancies'];
  const sheets = Object.fromEntries(names.map(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error('Не найден лист «' + name + '».');
    const range = sheet.getDataRange();
    // Inspector needs only the selected character, and no derived formulas are written.
    const values = name === 'Inspector' ? sheet.getRange('B1').getValues() : range.getValues();
    return [name, { sheet, values }];
  }));
  const { sheet, values: regValues } = sheets.Registration;
  if (!Number.isInteger(row) || row < 2 || row > regValues.length) {
    throw new Error('Строка регистрации недоступна.');
  }
  const headers = regValues[0];
  const values = regValues[row - 1];
  const timestampColumn = headers.indexOf('Отметка времени') + 1;
  if (!timestampColumn || !values[timestampColumn - 1] || headers[R.approved - 1] !== 'Статус' ||
      !R.charName || !R.gold || !R.inputLvl) {
    throw new Error('Проверьте строку и сопоставление колонок Registration.');
  }
  const name = String(getField(values, R, 'charName')).trim();
  const approved = isApproved(getField(values, R, 'approved'));
  const regFormulas = sheet.getDataRange().getFormulas();
  const notes = sheet.getRange(row, 1, 1, headers.length).getNotes()[0];
  const logs = sheets['LevelUp Log'].values.slice(1);
  const entries = logs.flatMap((values, index) => {
    const entryName = String(values[LOG.character - 1]).trim();
    return name && name.toLowerCase() === entryName.toLowerCase()
      ? [{ values, row: index + 2, name: entryName, level: Number(values[LOG.newLevel - 1]),
           timestamp: String(values[LOG.timestamp - 1]) }] : [];
  });
  const initial = values.slice();
  const blocked = new Set();
  rewindLevelCorrection_(entries.slice().sort((a, b) => b.level - a.level || b.row - a.row),
    initial, regFormulas[row - 1], name, blocked);
  const numberKeys = ['inputLvl', 'hpTotal', 'class1Lvl', 'class2Lvl', 'class3Lvl', 'fcHP',
    'fcSkill', 'fcRacial', 'str', 'dex', 'con', 'int', 'wis', 'cha', 'gold'];
  const numeric = new Set(numberKeys.map(key => R[key]));
  const fields = headers.slice(0, R.approved - 1).flatMap((label, index) => {
    const column = index + 1;
    if (!label || column === timestampColumn) return [];
    const isBlocked = blocked.has(column);
    const type = numeric.has(column) ? 'number'
      : /story|skills|features|spells|traits|feats|abilities|gm should know|drawbacks|anything else/i.test(label)
        ? 'multiline' : 'text';
    return [{ key: String(column), column, label: String(label).trim(), type,
      value: String((isBlocked ? values : initial)[index]), note: notes[index],
      formula: regFormulas[row - 1][index] || isBlocked,
      locked: isBlocked ? 'Поле изменено после повышения и расходится с журналом.' : '' }];
  });
  const dbValues = sheets.Database.values;
  const dbHeaders = dbValues[3] || [];
  const db = {};
  const charColumn = dbHeaders.indexOf('char');
  const dbRow = dbValues.slice(4).find(record => name && String(record[charColumn]).trim().toLowerCase() === name.toLowerCase());
  if (dbRow) dbHeaders.forEach((key, index) => { if (key && !(key in db)) db[key] = dbRow[index]; });
  const spellCatalog = loadSpellCatalog(ss);
  const revision = Utilities.base64EncodeWebSafe(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, JSON.stringify([row, columns, fields, regFormulas[row - 1],
      names.map(key => [sheets[key].sheet.getSheetId(), sheets[key].values]), spellCatalog])
  ));
  return { ss, sheet, sheets, row, name, approved, columns, values, regValues, regFormulas,
    headers, fields, entries, db, spellCatalog, revision };
}

function planRegistrationCorrection_(request, revision, input) {
  if (request.revision !== revision) {
    throw new Error('Данные изменились. Откройте панель заново и проверьте правку.');
  }
  const changes = request.fields.filter(field => Object.hasOwn(input, field.key))
    .map(field => parseGameCorrectionField_(field, input[field.key]))
    .filter(change => change.display !== change.field.value);
  if (!changes.length) throw new Error('Поля не изменились.');
  for (const change of changes) {
    if (change.field.formula) throw new Error(change.field.locked || 'Поле содержит формулу.');
  }
  const R = request.columns.registration;
  const changed = new Map(changes.map(change => [change.field.column, change]));
  const value = column => changed.get(column)?.value ?? request.values[column - 1];
  const name = String(value(R.charName)).trim();
  if (!name || /[,;:\r\n]|->|=>|[→>]/.test(name) ||
      (changed.has(R.charName) && request.regValues.some((values, index) =>
        index > 0 && index !== request.row - 1 && nameKey(getField(values, R, 'charName')) === nameKey(name)))) {
    throw new Error('Укажите уникальное короткое имя без разделителей форм.');
  }
  for (const column of [R.inputLvl, R.class1Lvl, R.class2Lvl, R.class3Lvl, R.hpTotal]) {
    if (changed.has(column) && (!Number.isInteger(Number(value(column))) ||
        (column === R.inputLvl && (Number(value(column)) < 1 || Number(value(column)) > 20)))) {
      throw new Error('Проверьте целое число в поле «' + changed.get(column).field.label + '».');
    }
  }

  const rename = Boolean(request.name && name !== request.name);
  const links = rename ? registrationRenameLinks_(request, name) : [];
  const warnings = [];
  const effects = [];
  const raceColumn = request.headers.indexOf('Race') + 1;
  const variantColumn = request.headers.indexOf('Variant Heritage (if any)') + 1;
  const constitutionChanged = [R.con, raceColumn, variantColumn].some(column => changed.has(column));
  const replayKeys = new Set([
    R.inputLvl, R.hpTotal, R.class1, R.class1Lvl, R.class1Arch, R.class2, R.class2Lvl, R.class2Arch,
    R.class3, R.class3Lvl, R.class3Arch, R.favoredClass, R.fcHP, R.fcSkill, R.fcRacial, R.con,
    R.skills, R.feats, R.classFeatures, R.specialAbil, R.traits, R.drawbacks, R.isCaster,
    R.cantNum, R.cant, raceColumn, variantColumn,
    ...[4, 8, 12, 16, 20].map(level => R['abilityInc' + level]),
    ...Array.from({ length: 18 }, (_, index) => R.spell1Num + index),
  ]);
  let replay = null;
  if (request.entries.length && changes.some(change => replayKeys.has(change.field.column))) {
    const db = { ...request.db };
    if (!db.char) throw new Error('В Database не найдена карточка для пересчёта повышений.');
    if (changed.has(R.con)) db.con_score = value(R.con);
    if (changed.has(raceColumn)) {
      const table = request.sheets['Reference tables'].values;
      const header = table.find(row => row[0] === 'Race' && row[3] === 'CON');
      const race = header && table.slice(table.indexOf(header) + 1)
        .find(row => row[0] === value(raceColumn));
      if (!race) throw new Error('Укажите расу как в справочнике Reference tables.');
      db.con_add = Number(race[3]) || 0;
    }
    replay = replayLevelCorrection_({ ...request, timestamp: '', form: request.sheets.LvlUP.values,
      database: [request.db] }, request.row, request.name, 1, null, {
      baseline: changes, database: db,
    });
    warnings.push('Будут пересчитаны уровни ' + replay.levels.join(', ') + '.');
    const replacedLogs = new Set(replay.oldLogs);
    for (let index = links.length - 1; index >= 0; index--) {
      if (links[index].sheet === request.sheets['LevelUp Log'].sheet && replacedLogs.has(links[index].row)) {
        links.splice(index, 1);
      }
    }
  }
  if (constitutionChanged && request.approved) {
    warnings.push('Проверьте исходные ОЗ и ручные расовые поправки в Database.');
  }
  if (rename) warnings.push('Короткое имя будет обновлено в связанных записях и списках выбора форм.');

  let newGold = null;
  const goldChanged = request.approved && changed.has(R.gold);
  if (goldChanged) {
    const txn = request.sheets.Transactions;
    const before = Number(getField(request.values, R, 'gold')) || 0;
    const after = Number(value(R.gold)) || 0;
    const entries = txn.values.slice(1).flatMap((values, index) =>
      values[1] === 'registration' && registrationNameMatches_(values[2], request.name) &&
      values[3] === 'gold' ? [{ values, row: index + 2 }] : []);
    if (entries.length > 1 || entries.some(entry => Number(entry.values[4]) !== before)) {
      throw new Error('Начальное золото расходится с Transactions. Проверьте журнал.');
    }
    if (entries.length) {
      registrationLink_(links, txn, entries[0].row, 5, after);
    } else if (after && !before) {
      newGold = after;
    } else if (before) {
      warnings.push('Начальное золото ещё не записано в Transactions. Проверьте Process Starting Gold после сохранения.');
    }
    if (entries.length || newGold !== null) {
      const delta = roundGold(after - before);
      if (delta) effects.push({ name, item: 'gold', delta });
      if (delta < 0) {
        let balance = 0;
        for (let index = 1; index < txn.values.length; index++) {
          const entry = txn.values[index];
          if (!registrationNameMatches_(entry[2], request.name) || entry[3] !== 'gold') continue;
          balance += Number(entry[4]) || 0;
          if (index + 1 > entries[0].row && entry[1] === 'transfer' && entry[4] < 0 &&
              balance + delta < -0.000001) {
            throw new Error('После правки не хватает золота для передачи в Transactions, строка ' + (index + 1) + '.');
          }
        }
      }
    }
  }
  const preview = {
    changes: changes.map(change => ({ label: change.field.label, from: change.field.value, to: change.display })),
    effects, warnings,
    registration: replay?.changes.map(change => ({
      label: String(change.label).trim(), from: String(change.before), to: String(change.after),
    })) || [],
  };
  return { changes, replay, links, name, rename, goldChanged, newGold, preview };
}

function registrationNameMatches_(value, name) {
  return name && String(value).trim().toLowerCase() === name.toLowerCase();
}

function registrationLink_(links, source, row, column, value) {
  if (source.sheet.getRange(row, column).getFormula()) {
    throw new Error('Связанная ячейка содержит формулу: ' + source.sheet.getName() +
      ', строка ' + row + ', колонка ' + column + '.');
  }
  links.push({ sheet: source.sheet, row, column, value });
}

function registrationRenameLinks_(request, name) {
  const old = request.name;
  const exact = value => registrationNameMatches_(value, old) ? name : value;
  const token = value => registrationNameMatches_(value, old)
    ? value.replace(/^(\s*).*?(\s*)$/s, (_match, before, after) => before + name + after) : value;
  const aliases = value => nameKey(value) === nameKey(old)
    ? value.replace(/^(\s*).*?(\s*)$/s, (_match, before, after) => before + name + after) : value;
  const list = value => String(value).split(',').map(token).join(',');
  const named = value => String(value).replace(/(^|;)([^:;]+)(?=:)/g,
    (_match, separator, character) => separator + token(character));
  const transfers = value => String(value).replace(/(^|[\n;])([^\n;]+?)(->|→|=>|>)([^:\n;]+)(?=:)/g,
    (_match, separator, from, arrow, to) => separator + aliases(from) + arrow + aliases(to));
  const transferNote = value => String(value).replace(/^(To |From )([^:]+)(?=:|$)/,
    (_match, prefix, character) => prefix + token(character));
  const columns = request.columns;
  const links = [];
  const scan = (sheetName, specs) => {
    const source = request.sheets[sheetName];
    if (specs.some(([column]) => !Number.isInteger(column) || column < 1)) {
      throw new Error('Не найдена связанная колонка на листе ' + sheetName + '.');
    }
    source.values.slice(1).forEach((values, index) => specs.forEach(([column, transform]) => {
      const raw = values[column - 1];
      if (raw === '' || raw === undefined) return;
      const value = transform(raw, values);
      if (value !== raw) registrationLink_(links, source, index + 2, column, value);
    }));
  };
  scan('LvlUP', [[columns.lvlup.charName, exact]]);
  scan('Games', [[columns.games.characters, list], [columns.games.goldDelta, named], [columns.games.loot, named]]);
  scan('Inventory(raw)', [[columns.inventory.charName, exact],
    [columns.inventory.participants, list], [columns.inventory.transfers, transfers]]);
  scan('Transactions', [[TXN.character, exact],
    [TXN.notes, (value, row) => row[TXN.source - 1] === 'transfer' ? transferNote(value) : value]]);
  scan('LevelUp Log', [[LOG.character, exact]]);
  const board = request.sheets['Bulletin Board'];
  scan('Bulletin Board', [[board.values[0].indexOf('Author') + 1, exact]]);
  for (const sheetName of ['Inventory', 'Feats', 'Spells', 'Discrepancies']) {
    const headers = request.sheets[sheetName].values[0];
    const index = headers.indexOf(sheetName === 'Discrepancies' ? 'Персонаж' : 'character');
    if (index >= 0) scan(sheetName, [[index + 1, exact]]);
  }
  const inspector = request.sheets.Inspector;
  if (registrationNameMatches_(inspector.values[0][0], old)) {
    registrationLink_(links, inspector, 1, 2, name);
  }
  return links;
}
