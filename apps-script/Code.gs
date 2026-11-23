/**
 * Анкета lovegames.kz — сервер на Google Apps Script (сайт анкеты на GitHub Pages шлёт сюда POST).
 * Хранит каждую анкету в Google Drive (папка на клиента: текст + фото)
 * и добавляет строку в таблицу «Анкеты lovegames.kz».
 */

// Почта для уведомлений о новой анкете. Оставьте пустой, если не нужно.
const NOTIFY_EMAIL = '';
// WhatsApp студии (для письма-уведомления)
const STUDIO_WA = '77064007235';

const ROOT_NAME = 'Анкеты lovegames.kz';
const HEAD = ['Дата', 'Контакт', 'Герои', 'Город', 'Повод', 'Нужна к', 'Сюрприз', 'Язык',
  'Папка', 'Фото (ожидается)', 'Фото (загружено)', 'Статус', 'Анкета'];

function doGet() {
  return ContentService.createTextOutput('lovegames ok');
}

/** Приём анкеты с сайта (GitHub Pages). */
function doPost(e) {
  let out;
  try {
    const d = JSON.parse(e.postData.contents);
    let res;
    if (d.action === 'start') res = startSubmission(d.meta || {});
    else if (d.action === 'photo') res = uploadPhoto(d.folderId, d.sub, d.name, d.data);
    else if (d.action === 'finish') res = finish(d.folderId, d.row, d.uploaded);
    else throw new Error('Неизвестное действие');
    out = { ok: true, res: res };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Запустите один раз вручную: создаст папку и таблицу и попросит доступы. */
function setup() {
  const s = store_();
  Logger.log('Папка: ' + s.folder.getUrl());
  Logger.log('Таблица: ' + s.sheet.getParent().getUrl());
}

function store_() {
  const p = PropertiesService.getScriptProperties();
  let folder = null, ss = null;
  try { folder = DriveApp.getFolderById(p.getProperty('FOLDER')); } catch (e) {}
  if (!folder) {
    folder = DriveApp.createFolder(ROOT_NAME);
    p.setProperty('FOLDER', folder.getId());
  }
  try { ss = SpreadsheetApp.openById(p.getProperty('SHEET')); } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.create(ROOT_NAME);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    const sh = ss.getSheets()[0];
    sh.setName('Анкеты');
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold');
    p.setProperty('SHEET', ss.getId());
  }
  return { folder: folder, sheet: ss.getSheetByName('Анкеты') || ss.getSheets()[0] };
}

function clean_(s, n) { return String(s || '').slice(0, n || 500); }

/** Шаг 1: создаёт папку клиента и строку в таблице. */
function startSubmission(meta) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = store_();
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm');
    const name = (stamp + ' ' + clean_(meta.heroes, 100)).replace(/[\\/:*?"<>|]/g, ' ');
    const f = s.folder.createFolder(name);
    f.createFile('Анкета.txt', clean_(meta.text, 200000), MimeType.PLAIN_TEXT);
    f.createFile('anketa.json', clean_(meta.json, 400000), 'application/json');
    s.sheet.appendRow([new Date(), clean_(meta.contact), clean_(meta.heroes), clean_(meta.city),
      clean_(meta.occasion), clean_(meta.deadline), clean_(meta.surprise), clean_(meta.lang),
      f.getUrl(), Number(meta.photoCount) || 0, 0, 'Загружаются фото', clean_(meta.text, 49000)]);
    return { folderId: f.getId(), row: s.sheet.getLastRow() };
  } finally {
    lock.releaseLock();
  }
}

function clientFolder_(folderId) {
  const root = store_().folder;
  const f = DriveApp.getFolderById(folderId);
  const parents = f.getParents();
  if (!parents.hasNext() || parents.next().getId() !== root.getId()) throw new Error('Неверная папка');
  return f;
}

/** Шаг 2: загружает одно фото (data URL, уже сжатое в браузере). */
function uploadPhoto(folderId, sub, fileName, dataUrl) {
  const f = clientFolder_(folderId);
  const m = /^data:(image\/[a-z+.-]+);base64,(.+)$/i.exec(String(dataUrl));
  if (!m) throw new Error('Это не картинка');
  const subName = clean_(sub, 80).replace(/[\\/:*?"<>|]/g, ' ') || 'Фото';
  const it = f.getFoldersByName(subName);
  const dir = it.hasNext() ? it.next() : f.createFolder(subName);
  dir.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], clean_(fileName, 120)));
  return true;
}

/** Шаг 3: отмечает анкету готовой и шлёт уведомление. */
function finish(folderId, row, uploaded) {
  const f = clientFolder_(folderId);
  const s = store_();
  row = Number(row);
  if (row > 1 && String(s.sheet.getRange(row, 9).getValue()).indexOf(folderId) !== -1) {
    s.sheet.getRange(row, 11).setValue(Number(uploaded) || 0);
    s.sheet.getRange(row, 12).setValue('Новая');
  }
  if (NOTIFY_EMAIL) {
    const heroes = row > 1 ? s.sheet.getRange(row, 3).getValue() : '';
    const contact = row > 1 ? s.sheet.getRange(row, 2).getValue() : '';
    MailApp.sendEmail(NOTIFY_EMAIL, 'Новая анкета: ' + heroes,
      'Герои: ' + heroes + '\nКонтакт: ' + contact + '\nФото: ' + uploaded +
      '\nПапка: ' + f.getUrl() + '\nТаблица: ' + s.sheet.getParent().getUrl());
  }
  return true;
}
