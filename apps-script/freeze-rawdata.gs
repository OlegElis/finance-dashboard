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
 * ⚠ Дашборд читает колонки по ПОЗИЦИЯМ (COL в index.html). Скрипт копирует диапазон как есть и порядок
 *   не меняет - но если поменяется порядок колонок в источнике, сломаются отчёты, а не скрипт.
 */

var SRC_SHEET = 'RawData_src'; // лист с формулами (IMPORTRANGE)
var DST_SHEET = 'RawData';     // лист-значения, его читает дашборд
var HEAD_COLS = 15;            // столько колонок в шапке ждёт дашборд
var MIN_ROWS_RATIO = 0.5;      // не затирать, если строк вдруг стало меньше половины от прошлого раза

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

  // ── Запись. Одним setValues (батчем), а не построчно: 60 тыс. строк построчно идут минутами.
  //    clearContents, а НЕ clear(): форматы и ширины колонок листа сохраняем.
  dst.clearContents();
  dst.getRange(1, 1, vals.length, vals[0].length).setValues(vals);

  var p = PropertiesService.getDocumentProperties();
  p.setProperty('lastRows', String(vals.length));
  p.setProperty('lastFreezeAt', new Date().toISOString());
  Logger.log('RawData залит значениями: ' + vals.length + ' строк');
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
