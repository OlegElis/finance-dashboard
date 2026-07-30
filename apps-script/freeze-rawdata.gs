/**
 * ЗАЛИВКА RawData ЗНАЧЕНИЯМИ (вместо живых формул IMPORTRANGE).
 *
 * Зачем: замер 30.07.2026 показал, что дашборд ждёт не данные, а ПЕРЕСЧЁТ формул на стороне Google.
 *   холодный запрос ОДНОЙ ячейки формульного листа - 14 894 мс
 *   тот же запрос по прогретому листу        -    376 мс
 *   весь лист (23,5 МБ, 60 700 строк)        -  2 539 мс ответ + 1 266 мс тело, разбор 40 мс
 * То есть объём данных ни при чём: платим за холодный пересчёт IMPORTRANGE, и платит его КАЖДЫЙ
 * пользователь при открытии. Если лист лежит значениями, пересчитывать нечего - остаётся только выкачка.
 *
 * Схема: формулы переезжают на отдельный лист-источник (SRC), дашборд читает лист-значения (DST).
 *   RawData_src - IMPORTRANGE/QUERY, как сейчас в RawData; человеку не нужен, можно скрыть
 *   RawData     - плоские значения, их пишет этот скрипт; ЕГО и читает дашборд (COL по позициям)
 *
 * Установка:
 *   1) Переименовать текущий формульный RawData → RawData_src, скрыть его.
 *   2) Создать пустой лист RawData (именно с таким именем - дашборд ходит за ним).
 *   3) Расширения → Apps Script, вставить этот файл, сохранить.
 *   4) Один раз выполнить freezeRawData() вручную и убедиться, что DST заполнился.
 *   5) Триггеры → добавить триггер: freezeRawData, по времени, раз в час (или чаще/реже по нужде).
 *
 * Первые два-три прогона проверить глазами: в RawData лежат ЗНАЧЕНИЯ, а не формулы (клик по ячейке -
 * в строке формул само число); число строк совпало с источником; итоги на дашборде не поехали.
 * Снимки предыдущих значений лежат скрытыми листами RawData_bak_ГГГГ-ММ-ДД_ЧЧММ (последние три).
 *
 * ⚠ Дашборд читает колонки по ПОЗИЦИЯМ (COL в index.html). Скрипт копирует диапазон как есть и порядок
 *   не меняет - но если поменяется порядок колонок в источнике, сломаются отчёты, а не скрипт.
 */

var SRC_SHEET = 'RawData_src'; // лист с формулами (IMPORTRANGE)
var DST_SHEET = 'RawData';     // лист-значения, его читает дашборд
var HEAD_COLS = 15;            // столько колонок в шапке ждёт дашборд
var MIN_ROWS_RATIO = 0.5;      // не затирать, если строк вдруг стало меньше половины от прошлого раза
var SNAP_PREFIX = 'RawData_bak_'; // датированные снимки предыдущих значений
var SNAP_KEEP = 3;                // сколько снимков держать

function freezeRawData() {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName(SRC_SHEET);
  var dst = ss.getSheetByName(DST_SHEET);
  if (!src) throw new Error('Нет листа-источника ' + SRC_SHEET);
  if (!dst) throw new Error('Нет листа-приёмника ' + DST_SHEET);

  // SpreadsheetApp.flush() заставляет досчитать формулы ДО чтения: иначе можно вычитать транзитное состояние
  SpreadsheetApp.flush();
  var vals = src.getDataRange().getValues();

  // ── Страховки. Ни одна из них не «перестраховка»: дашборд читает лист вслепую по позициям,
  //    и однажды записанный мусор он покажет как валидные данные.
  if (!vals.length) return abort_('источник пуст');

  // 1) Транзитное состояние IMPORTRANGE: шапка приходит пустой / #ERROR! / Loading…
  //    Ровно от этого в дашборде стоит проверка структуры листа - фиксировать такое состояние нельзя.
  var head = vals[0].slice(0, HEAD_COLS).map(function (h) { return String(h == null ? '' : h).trim().toLowerCase(); });
  var headOk = head.length === HEAD_COLS && head.every(function (h) {
    return h !== '' && h.charAt(0) !== '#' && h.indexOf('loading') < 0 && h.indexOf('n/a') < 0;
  });
  if (!headOk) return abort_('шапка невалидна (идёт пересчёт импорта?): ' + JSON.stringify(head));

  // 2) Обвал числа строк - признак сбоя у источника, а не удаления данных
  var prev = Number(PropertiesService.getDocumentProperties().getProperty('lastRows') || 0);
  if (prev && vals.length < prev * MIN_ROWS_RATIO) {
    return abort_('строк стало ' + vals.length + ' против ' + prev + ' в прошлый раз - похоже на сбой источника');
  }

  // 3) Снимок ПРЕДЫДУЩИХ значений перед перезаписью. Скрипт пишет в источник истины, и правило
  //    «вчерашние верные лучше сегодняшних битых» работает, только если вчерашние где-то лежат.
  //    Ловятся случаи, которые проверки 1-2 пропускают: источник отдал валидную по форме, но неверную
  //    по сути выгрузку (сдвиг колонок, чужой период, половина направлений).
  snapshot_(ss, dst);

  // ── Запись. Одним setValues (батчем), а не построчно: 60 тыс. строк построчно идут минутами.
  //    clearContents, а НЕ clear(): форматы и ширины колонок листа сохраняем.
  dst.clearContents();
  dst.getRange(1, 1, vals.length, vals[0].length).setValues(vals);

  var p = PropertiesService.getDocumentProperties();
  p.setProperty('lastRows', String(vals.length));
  p.setProperty('lastFreezeAt', new Date().toISOString());
  Logger.log('RawData залит значениями: ' + vals.length + ' строк');
}

/**
 * Снимок текущего DST в датированный лист + подчистка старых (держим SNAP_KEEP штук).
 * copyTo, а не getValues/setValues: копирование листа целиком идёт одной операцией, без прогона
 * 900 тыс. ячеек через скрипт (это минуты и риск упереться в лимит времени выполнения).
 * ⚠ ЛИМИТ КНИГИ - 10 млн ячеек. Сейчас 60 700 строк × 15 = 911 тыс. ячеек на копию:
 *   источник + рабочий лист + 3 снимка ≈ 4,6 млн, запас есть. Но при удвоении базы (≈120 тыс. строк)
 *   те же пять листов дадут ≈9 млн и упрутся в потолок. Тогда либо уменьшить SNAP_KEEP, либо
 *   держать снимки в ОТДЕЛЬНОЙ книге / выгружать CSV в Drive.
 */
function snapshot_(ss, dst) {
  if (dst.getLastRow() < 2) return; // первый запуск: снимать нечего
  var tz = ss.getSpreadsheetTimeZone() || 'Europe/Moscow';
  var name = SNAP_PREFIX + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HHmm');
  var old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old); // повторный запуск в ту же минуту - перезаписываем снимок, а не плодим «Copy of»
  var snap = dst.copyTo(ss).setName(name);
  snap.hideSheet(); // снимки человеку в глаза не лезут

  // Подчистка: имя содержит дату в ISO-порядке, поэтому лексикографическая сортировка = хронологическая.
  var all = ss.getSheets()
    .map(function (sh) { return sh.getName(); })
    .filter(function (n) { return n.indexOf(SNAP_PREFIX) === 0; })
    .sort();
  while (all.length > SNAP_KEEP) {
    var drop = all.shift();
    ss.deleteSheet(ss.getSheetByName(drop));
    Logger.log('Снимок удалён (старше ' + SNAP_KEEP + '): ' + drop);
  }
}

/** Отказ от записи: прошлые данные ОСТАЮТСЯ на месте - лучше вчерашние верные, чем сегодняшние битые. */
function abort_(why) {
  Logger.log('ОТКАЗ от заливки: ' + why);
  PropertiesService.getDocumentProperties().setProperty('lastAbort', new Date().toISOString() + ' · ' + why);
  return null;
}

/** Пункт меню, чтобы гонять вручную, не открывая редактор скриптов. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Дашборд')
    .addItem('Залить RawData значениями', 'freezeRawData')
    .addToUi();
}
