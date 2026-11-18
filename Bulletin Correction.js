function showBulletinCorrection() {
  const ui = SpreadsheetApp.getUi();
  const range = SpreadsheetApp.getActiveRange();
  if (!range || range.getSheet().getName() !== 'Bulletin Board' ||
      range.getRow() < 2 || range.getNumRows() !== 1) {
    ui.alert('Выберите одну строку объявления на листе Bulletin Board.');
    return;
  }
  try {
    const bulletin = readBulletinCorrection_(range.getRow());
    const page = HtmlService.createTemplateFromFile('Game Reward Sidebar');
    page.initial = {
      kind: 'bulletin',
      row: bulletin.row,
      revision: bulletin.revision,
      fields: bulletin.fields,
      context: 'Bulletin Board · строка ' + bulletin.row + ' · ' +
        (bulletin.status || 'без статуса'),
    };
    ui.showSidebar(page.evaluate().setTitle('Исправить объявление'));
  } catch (error) {
    ui.alert(error.message);
  }
}

function previewBulletinCorrection(row, revision, input) {
  const bulletin = readBulletinCorrection_(row);
  const changes = planBulletinCorrection_(bulletin, revision, input);
  const publicChange = changes.some(change => change.field.key !== 'gm');
  let publication = 'Изменятся только сведения для мастера.';
  if (publicChange) {
    publication = bulletin.status === 'Утверждено'
      ? 'Изменения появятся в списке утверждённых объявлений.'
      : 'Объявление появится в списке после утверждения.';
  }
  return {
    changes: changes.map(change => ({
      label: change.field.label, from: change.field.value, to: change.display,
    })),
    effects: [],
    publication,
  };
}

function applyBulletinCorrection(row, revision, input) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const bulletin = readBulletinCorrection_(row);
    const changes = planBulletinCorrection_(bulletin, revision, input);
    const date = Utilities.formatDate(
      new Date(), bulletin.ss.getSpreadsheetTimeZone(), 'dd.MM.yyyy HH:mm:ss'
    );
    const requests = changes.map(change => {
      const entry = date + ' — исправлено мастером: ' +
        (change.field.value || 'пусто') + ' → ' + (change.display || 'пусто') + '.';
      return gameCellsUpdate_(bulletin.sheet, row, change.field.column, [[{
        userEnteredValue: gameCellValue_(change.value),
        note: change.field.note ? change.field.note + '\n' + entry : entry,
      }]], 'userEnteredValue,note');
    });
    Sheets.Spreadsheets.batchUpdate({ requests }, bulletin.ss.getId());
    return 'Правка сохранена. История добавлена в примечания к изменённым полям.';
  } finally {
    lock.releaseLock();
  }
}

function readBulletinCorrection_(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Bulletin Board');
  if (!sheet || !Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) {
    throw new Error('Строка объявления недоступна.');
  }
  const count = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, count).getValues()[0];
  const required = ['Отметка времени', 'Author', 'Title', 'Announcement',
    'Is there anything GM should know?', 'Статус'];
  if (required.some(header => headers.filter(value => value === header).length !== 1)) {
    throw new Error('Проверьте заголовки листа Bulletin Board.');
  }
  const range = sheet.getRange(row, 1, 1, count);
  const values = range.getValues()[0];
  const notes = range.getNotes()[0];
  const formulas = range.getFormulas()[0];
  const timestamp = values[headers.indexOf('Отметка времени')];
  if (!timestamp) throw new Error('В строке нет ответа формы.');
  const status = String(values[headers.indexOf('Статус')]);
  const fields = [
    ['author', 'Author', 'Автор', 'text'],
    ['title', 'Title', 'Заголовок', 'text'],
    ['announcement', 'Announcement', 'Объявление', 'multiline'],
    ['gm', 'Is there anything GM should know?', 'Сведения для мастера', 'multiline'],
  ].map(([key, header, label, type]) => {
    const index = headers.indexOf(header);
    return { key, label, type, column: index + 1, value: String(values[index]),
      note: notes[index], formula: formulas[index] };
  });
  const revision = Utilities.base64EncodeWebSafe(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify([sheet.getSheetId(), row, timestamp, status, fields])
  ));
  return { ss, sheet, row, fields, revision, status };
}

function planBulletinCorrection_(bulletin, revision, input) {
  if (bulletin.revision !== revision) {
    throw new Error('Объявление изменилось. Откройте панель заново и проверьте правку.');
  }
  const changes = bulletin.fields.filter(field => Object.hasOwn(input, field.key))
    .map(field => {
      const change = parseGameCorrectionField_(field, input[field.key]);
      if (change.display !== field.value) {
        if (field.formula) throw new Error('Поле содержит формулу: ' + field.label);
        if (field.key !== 'gm' && !change.value) {
          throw new Error('Заполните поле «' + field.label + '».');
        }
        if (field.key === 'author' && !loadCharacterNames(bulletin.ss).includes(change.value)) {
          throw new Error('Укажите имя персонажа как на листе Database.');
        }
      }
      return change;
    }).filter(change => change.display !== change.field.value);
  if (!changes.length) throw new Error('Поля не изменились.');
  return changes;
}
