function showLevelUpCorrection() {
  const ui = SpreadsheetApp.getUi();
  const range = SpreadsheetApp.getActiveRange();
  if (!range || range.getSheet().getName() !== 'LvlUP' ||
      range.getRow() < 2 || range.getNumRows() !== 1) {
    ui.alert('Выберите строку повышения на листе LvlUP.');
    return;
  }
  try {
    const request = readLevelUpCorrection_(range.getRow());
    const page = HtmlService.createTemplateFromFile('Game Reward Sidebar');
    page.initial = { kind: 'levelup', row: request.row, revision: request.revision,
      fields: request.fields, context: 'LvlUP · строка ' + request.row + ' · ' +
        (request.selectedEntries.length ? 'применено' : 'ещё не применено') };
    ui.showSidebar(page.evaluate().setTitle('Исправить повышение уровня'));
  } catch (error) {
    ui.alert(error.message);
  }
}

function previewLevelUpCorrection(row, revision, input) {
  return planLevelUpCorrection_(readLevelUpCorrection_(row), revision, input).preview;
}

function applyLevelUpCorrection(row, revision, input) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const request = readLevelUpCorrection_(row);
    const plan = planLevelUpCorrection_(request, revision, input);
    const now = new Date();
    const date = Utilities.formatDate(now, request.ss.getSpreadsheetTimeZone(), 'dd.MM.yyyy HH:mm:ss');
    const updates = plan.changes.map(change => {
      const entry = date + ' — исправлено мастером: ' + (change.field.value || 'пусто') +
        ' → ' + (change.display || 'пусто') + '.';
      return gameCellsUpdate_(request.sheet, row, change.field.column, [[{
        userEnteredValue: gameCellValue_(change.value),
        note: change.field.note ? change.field.note + '\n' + entry : entry,
      }]], 'userEnteredValue,note');
    });

    plan.characters.forEach(character => {
      character.changes.forEach(change => {
        updates.push(gameCellsUpdate_(request.registration, character.row, change.column,
          [[{ userEnteredValue: gameCellValue_(change.after) }]], 'userEnteredValue'));
      });
    });

    if (plan.oldLogs.length) {
      const sheetId = request.log.getSheetId();
      plan.oldLogs.slice().sort((a, b) => b - a).forEach(logRow => {
        updates.push({ deleteDimension: { range: {
          sheetId, dimension: 'ROWS', startIndex: logRow - 1, endIndex: logRow,
        } } });
      });
      const logRows = plan.characters.flatMap(character => character.events.flatMap(event => {
        const changes = event.changes.length ? event.changes : [{ field: '(no changes)', regCol: 0, old: '', new: '' }];
        return changes.map((change, index) => [
          event.timestamp, now.toISOString(), character.name, event.level, change.field,
          change.regCol, change.old, change.new,
          index === 0 ? String(getField(event.values, request.columns.lvlup, 'gmNotes')).trim() : '',
        ]);
      }));
      if (logRows.length) {
        updates.push({ appendCells: { sheetId, fields: 'userEnteredValue',
          rows: logRows.map(values => ({ values: values.map(value => ({
            userEnteredValue: gameCellValue_(value),
          })) })),
        } });
      }
    }

    Sheets.Spreadsheets.batchUpdate({ requests: updates }, request.ss.getId());
    if (plan.characters.some(character => character.changes.length)) {
      try {
        SpreadsheetApp.flush();
        parseFeats();
        parseSkills();
        parseSpells();
        SpreadsheetApp.flush();
      } catch (error) {
        console.error(error);
        return 'Правка сохранена. Обновите списки через PF1e Tools → Refresh ALL normalized tables.';
      }
    }
    return 'Правка сохранена. История добавлена в примечания к изменённым полям.';
  } finally {
    lock.releaseLock();
  }
}

function readLevelUpCorrection_(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('LvlUP');
  const registration = ss.getSheetByName('Registration');
  const log = ss.getSheetByName('LevelUp Log');
  if (!sheet || !registration || !Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) {
    throw new Error('Строка повышения недоступна.');
  }
  const columns = loadColumnMappings(ss);
  const L = columns.lvlup;
  const R = columns.registration;
  const form = sheet.getDataRange().getValues();
  const values = form[row - 1];
  const headers = form[0];
  if (![L.timestamp, L.charName, L.approved, L.newLevel, L.hpRolled].every(column =>
    Number.isInteger(column) && column > 0 && column <= headers.length) ||
      headers[L.approved - 1] !== 'Статус' || !R.charName || !R.inputLvl) {
    throw new Error('Проверьте сопоставление колонок LvlUP и Registration.');
  }
  const range = sheet.getRange(row, 1, 1, headers.length);
  const notes = range.getNotes()[0];
  const formulas = range.getFormulas()[0];
  const timestamp = String(getField(values, L, 'timestamp'));
  if (!timestamp || form.slice(1).filter(values => String(getField(values, L, 'timestamp')) === timestamp).length !== 1) {
    throw new Error('Для исправления нужна уникальная отметка времени ответа.');
  }
  const fields = headers.slice(0, L.approved - 1).flatMap((label, index) => {
    const column = index + 1;
    if (!label || column === L.timestamp) return [];
    const type = [L.newLevel, L.hpRolled].includes(column) ? 'number'
      : column >= (L.newFeats || L.approved) ? 'multiline' : 'text';
    return [{ key: String(column), label, type, column, value: String(values[index]),
      note: notes[index], formula: formulas[index] }];
  });
  const logs = log && log.getLastRow() > 1
    ? log.getRange(2, 1, log.getLastRow() - 1, LOG.gmNotes).getValues() : [];
  const entries = logs.flatMap((values, index) => {
    const name = String(values[LOG.character - 1]).trim();
    return name ? [{ values, row: index + 2, name,
      timestamp: String(values[LOG.timestamp - 1]), level: Number(values[LOG.newLevel - 1]) }] : [];
  });
  const selectedEntries = entries.filter(entry => entry.timestamp === timestamp);
  const regRange = registration.getDataRange();
  const regValues = regRange.getValues();
  const regFormulas = regRange.getFormulas();
  const spellCatalog = loadSpellCatalog(ss);
  const db = ss.getSheetByName('Database');
  const dbData = db ? db.getDataRange().getValues() : [];
  const dbKeys = dbData[3] || [];
  const dbColumns = ['char', 'con_score', 'con_add', 'con_racial_var', 'con_ability_increase', 'con_mod']
    .map(key => [key, dbKeys.indexOf(key)]);
  const database = dbData.slice(4).map(row =>
    Object.fromEntries(dbColumns.map(([key, index]) => [key, index < 0 ? '' : row[index]])));
  const revision = Utilities.base64EncodeWebSafe(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify([sheet.getSheetId(), registration.getSheetId(), log?.getSheetId(), row,
      columns, form, fields, logs, regValues, regFormulas, spellCatalog, database])
  ));
  return { ss, sheet, registration, log, row, timestamp, values, form, columns, entries,
    selectedEntries, regValues, regFormulas, spellCatalog, database, fields, revision };
}

function planLevelUpCorrection_(request, revision, input) {
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
  const L = request.columns.lvlup;
  const charBefore = String(getField(request.values, L, 'charName')).trim();
  const charAfter = String(getField(values, L, 'charName')).trim();
  const oldLevel = Number(getField(request.values, L, 'newLevel'));
  const level = Number(getField(values, L, 'newLevel'));
  const roll = Number(getField(values, L, 'hpRolled'));
  if (!charAfter || !Number.isInteger(level) || level < 2 ||
      !Number.isInteger(roll) || roll < 0) {
    throw new Error('Укажите персонажа, целый уровень от 2 и целое число выпавших хитов.');
  }
  const targetRow = levelCorrectionCharacterRow_(request, charAfter);
  const preview = {
    changes: changes.map(change => ({ label: change.field.label,
      from: change.field.value, to: change.display })),
    effects: [], warnings: [],
  };
  if (!request.selectedEntries.length) {
    preview.publication = 'Правка вступит в силу при обработке повышения.';
    return { changes, characters: [], oldLogs: [], preview };
  }

  const original = levelCorrectionCharacterRow_(request, charBefore);
  if (!isApproved(getField(request.values, L, 'approved')) ||
      request.selectedEntries.some(entry => entry.name.toLowerCase() !== charBefore.toLowerCase() ||
        entry.level !== oldLevel)) {
    throw new Error('Статус или данные заявки расходятся с LevelUp Log. Проверьте журнал.');
  }
  const corrected = { timestamp: request.timestamp, level, values };
  const same = original === targetRow;
  const characters = same
    ? [replayLevelCorrection_(request, original, charBefore, Math.min(oldLevel, level), corrected)]
    : [replayLevelCorrection_(request, original, charBefore, oldLevel, null),
       replayLevelCorrection_(request, targetRow, charAfter, level, corrected)];
  const oldLogs = characters.flatMap(character => character.oldLogs);
  preview.warnings = characters.map(character => character.name + ' · пересчёт уровней ' +
    character.levels.join(', '));
  preview.registration = characters.flatMap(character => character.changes.map(change => ({
    label: character.name + ' · ' + change.label, from: String(change.before), to: String(change.after),
  })));
  return { changes, characters, oldLogs, preview };
}

function levelCorrectionCharacterRow_(request, name) {
  const R = request.columns.registration;
  const matches = request.regValues.flatMap((values, index) =>
    index > 0 && String(getField(values, R, 'charName')).trim().toLowerCase() === name.toLowerCase()
      ? [index + 1] : []);
  if (matches.length !== 1) throw new Error('В Registration нужна одна запись персонажа «' + name + '».');
  if (R.approved && !isApproved(getField(request.regValues[matches[0] - 1], R, 'approved'))) {
    throw new Error('Регистрация персонажа «' + name + '» должна быть утверждена.');
  }
  return matches[0];
}

function replayLevelCorrection_(request, row, name, firstLevel, corrected, options) {
  const L = request.columns.lvlup;
  const R = request.columns.registration;
  const original = request.regValues[row - 1];
  const formulas = request.regFormulas[row - 1];
  const staged = original.slice();
  const tail = request.entries.filter(entry => entry.name.toLowerCase() === name.toLowerCase() &&
    entry.level >= firstLevel);
  const byTimestamp = new Map();
  tail.forEach(entry => {
    const group = byTimestamp.get(entry.timestamp) || [];
    group.push(entry);
    byTimestamp.set(entry.timestamp, group);
  });

  const groups = [...byTimestamp.values()].sort((a, b) => a[0].level - b[0].level);
  if (new Set(groups.map(group => group[0].level)).size !== groups.length) {
    throw new Error('В журнале несколько применённых ответов на один уровень у ' + name + '.');
  }
  const events = groups.flatMap(group => {
    const { timestamp, level } = group[0];
    const rows = request.form.slice(1).filter(values => String(getField(values, L, 'timestamp')) === timestamp);
    if (rows.length !== 1 || Number(getField(rows[0], L, 'newLevel')) !== level ||
        String(getField(rows[0], L, 'charName')).trim().toLowerCase() !== name.toLowerCase() ||
        !isApproved(getField(rows[0], L, 'approved'))) {
      throw new Error('Уровень ' + level + ' у ' + name + ' расходится с заявкой или статусом.');
    }
    return timestamp === request.timestamp ? [] : [{ timestamp, level, values: rows[0] }];
  });

  rewindLevelCorrection_(groups.slice().reverse().flatMap(group => group.slice().reverse()),
    staged, formulas, name);
  const originalLevelBefore = Number(getField(staged, R, 'inputLvl'));
  if (groups.length && groups[0][0].level !== originalLevelBefore + 1) {
    throw new Error('У ' + name + ' не сходится уровень перед исправляемой цепочкой.');
  }
  options?.baseline?.forEach(change => { staged[change.field.column - 1] = change.value; });

  if (corrected) events.push(corrected);
  events.sort((a, b) => a.level - b.level);
  const adapter = {
    getRange: function(_row, column) {
      const index = column - 1;
      if (!Number.isInteger(column) || index < 0 || index >= staged.length || formulas[index]) {
        throw new Error('В Registration недоступна колонка ' + column + '.');
      }
      return {
        getValue: () => staged[index],
        setValue: value => { staged[index] = value; },
      };
    },
  };
  const db = options?.database ||
    request.database.find(entry => String(entry.char).trim().toLowerCase() === name.toLowerCase());
  const readDatabaseRow = () => {
    if (!db) throw new Error('В Database не найдены характеристики персонажа «' + name + '».');
    const milestones = [4, 8, 12, 16, 20].map(level => 'abilityInc' + level);
    const count = values => milestones.filter(key =>
      String(getField(values, R, key)).trim().toUpperCase() === 'CON').length;
    const increase = (Number(db.con_ability_increase) || 0) - count(original) + count(staged);
    const score = (Number(db.con_score) || 0) + (Number(db.con_add) || 0) +
      (Number(db.con_racial_var) || 0) + increase;
    return { ...db, con_ability_increase: increase, con_mod: Math.floor((score - 10) / 2) };
  };
  const levelBefore = Number(getField(staged, R, 'inputLvl'));
  if (!Number.isInteger(levelBefore) || levelBefore < 1) {
    throw new Error('У ' + name + ' не сходится уровень перед исправляемой цепочкой.');
  }

  events.forEach((event, index) => {
    if (event.level !== levelBefore + index + 1) {
      throw new Error('У ' + name + ' нарушен порядок повышений перед уровнем ' + event.level + '.');
    }
    event.changes = applyLevelUp(adapter, event.values, row, L, R, request.spellCatalog, readDatabaseRow);
    const className = String(getField(event.values, L, 'classToLevel')).trim();
    if ((className && !event.changes.some(change => /^class[123]Lvl$/.test(change.field))) ||
        event.changes.some(change => !change.regCol && /^(IGNORED|NOT FOUND)/.test(change.new))) {
      throw new Error('Проверьте класс, характеристику или переучивание заклинаний на уровне ' +
        event.level + ' у ' + name + '.');
    }
  });
  const changes = staged.flatMap((after, index) => String(after) !== String(original[index])
    ? [{ column: index + 1, before: original[index], after,
         label: request.regValues[0][index] || 'колонка ' + (index + 1) }] : []);
  const levels = [...new Set([...groups.map(group => group[0].level), ...events.map(event => event.level)])];
  return { row, name, events, changes, levels, oldLogs: tail.map(entry => entry.row) };
}

function rewindLevelCorrection_(entries, staged, formulas, name, blocked) {
  for (const entry of entries) {
    const field = entry.values[LOG.field - 1];
    const column = Number(entry.values[LOG.regCol - 1]);
    if (!column || blocked?.has(column)) continue;
    const index = column - 1;
    if (!Number.isInteger(column) || index < 0 || index >= staged.length || formulas[index] ||
        String(staged[index]) !== String(entry.values[LOG.newValue - 1])) {
      if (blocked) {
        blocked.add(column);
        continue;
      }
      throw new Error('Поле «' + field + '» уровня ' + entry.level + ' у ' + name +
        ' сейчас отличается от журнала или содержит формулу.');
    }
    const previous = entry.values[LOG.oldValue - 1];
    staged[index] = previous !== '' && /^(input_lvl|hpTotal|class[123]Lvl|fcHP|fcSkill|fcRacial)$/.test(field)
      ? Number(previous) : previous;
  }
}
