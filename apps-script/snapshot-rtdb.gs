/**
 * ПУБЛИКАЦИЯ СНИМКА ДАННЫХ В FIREBASE RTDB (заход 1 из 2 - только запись).
 *
 * Зачем: сейчас дашборд при каждом открытии тянет из Sheets API 23,5 МБ и ждёт пересчёта
 * IMPORTRANGE (замер 30.07.2026: холодный старт ~15 с, после заливки значениями ~3,6 с, и оба
 * растут линейно по числу строк). Снимок - ТОТ ЖЕ массив, подготовленный заранее, сжатый gzip
 * (~23,5 МБ → ~3 МБ, в base64 ~4 МБ) и положенный в RTDB. Расчётный слой дашборда не меняется.
 *
 * ⚠ ЭТОТ ЗАХОД НИЧЕГО НЕ ПЕРЕКЛЮЧАЕТ. Скрипт только пишет ветку snapshots; дашборд про неё ещё
 *   не знает и продолжает читать Sheets API. Смысл разделения - посмотреть на снимок и убедиться,
 *   что в нём ровно те же данные, ДО того как что-то переключать. Чтение - следующим заданием.
 *
 * Источник - МАСТЕР НАПРЯМУЮ (SpreadsheetApp.openById по правам владельца скрипта), а не лист
 * 00.UnifiedData через IMPORTRANGE. Это главная выгода помимо скорости: в пути публикации нет ни
 * восьми блоков импорта, ни потолка в 80 000 строк, ни гонки с пересборкой сводного листа.
 * ⚠ Колонки берутся ПОЗИЦИЯМИ, ровно A:O, порядок сохраняется: дашборд адресует их через COL
 *   (index.html:1082), перестановка колонок молча испортит ВСЕ отчёты, а не уронит скрипт.
 *
 * ─────────────────────────────────────────────────────────────────────────────────────────────
 * УСТАНОВКА
 *   1) Расширения → Apps Script (или отдельный standalone-проект - скрипт ходит в мастер по ID
 *      и container-bound ему быть не обязательно). Вставить этот файл, сохранить.
 *   2) Настройки проекта → «Показать файл манифеста appsscript.json» → добавить в него блок
 *      oauthScopes ЦЕЛИКОМ (без них ScriptApp.getOAuthToken() отдаст токен без прав на RTDB,
 *      и запись вернёт 401 - без внятного текста):
 *
 *        "oauthScopes": [
 *          "https://www.googleapis.com/auth/spreadsheets",
 *          "https://www.googleapis.com/auth/firebase.database",
 *          "https://www.googleapis.com/auth/userinfo.email",
 *          "https://www.googleapis.com/auth/script.external_request",
 *          "https://www.googleapis.com/auth/script.scriptapp",
 *          "https://www.googleapis.com/auth/script.send_mail"
 *        ]
 *
 *   3) Первая авторизация: выбрать функцию publishSnapshot и запустить РУКАМИ, принять запрос прав.
 *      Из триггера первый раз запускать нельзя - диалог авторизации показать некому.
 *   4) Триггеры → добавить: publishSnapshot, по времени, каждые 15 минут. Публикация по факту
 *      изменения мастера - отдельным заданием, после того как заработает чтение.
 *
 * ⚠ ПРАВА В RTDB. Токен ScriptApp принадлежит владельцу скрипта; если он же участник проекта
 *   Firebase, запись идёт МИМО правил базы (так же, как из консоли) - отдельного служебного
 *   аккаунта и ключа не нужно, и в исходнике нет ни одного секрета. Из этого следует две вещи:
 *   (а) правило записи на snapshots скрипту не требуется, но поставить его всё равно надо, чтобы
 *       ветку не мог переписать посторонний вошедший:  "snapshots": { ".read": "auth != null",
 *       ".write": "auth.token.email === 'o017ev@gmail.com'" };
 *   (б) ".read" тут заведено ЗАРАНЕЕ, под заход 2 - без него дашборд снимок не прочитает.
 *
 * ЧТО ПРОВЕРИТЬ ПОСЛЕ ПЕРВОГО ПРОГОНА
 *   - snapshotStatus() печатает в журнал meta, и meta.rows совпадает с числом строк листа
 *     00.UnifiedData (шапка входит в счёт);
 *   - в консоли Firebase видна ветка snapshots/rawdata с ОДНИМ поколением;
 *   - meta.bytes - порядка 3-4 МБ (это размер base64, то есть то, что реально скачает клиент;
 *     сам gzip примерно на четверть меньше);
 *   - повторный запуск publishSnapshot() без правок в мастере НИЧЕГО не публикует и пишет
 *     в журнал «данные не менялись» (см. п. «контрольная сумма» ниже).
 *
 * КОНТРОЛЬНАЯ СУММА - MD5 ВСЕГО МАССИВА (sum_), а не число строк и суммы денег.
 * ⚠ Так было до 16.08.2026, и это был немой отказ: сумма считалась как «строк + Σ поступлений +
 *   Σ списаний», поэтому правка любой НЕЧИСЛОВОЙ ячейки (ID_Offer, статья, контрагент, назначение,
 *   счёт, направление, ФНС) не меняла ни одного слагаемого - скрипт уходил в ветку «данные не
 *   менялись» и не публиковал НИКОГДА, пока не сдвинутся деньги или число строк. Поймано на
 *   практике 15.08.2026: правка ID_Offer в двух строках до дашборда не доехала.
 *   Теперь сверка идёт по MD5 от того самого JSON, который и уезжает в базу, - меняется что
 *   угодно в любой ячейке, меняется и сумма. Число строк и суммы денег в ней остались префиксом:
 *   по MD5 не видно НИЧЕГО, а по префиксу в журнале сразу читается масштаб изменения.
 * ⚠ ФОРМАТ 'sum' СМЕНИЛСЯ, поэтому ПЕРВЫЙ прогон после выкладки опубликует снимок независимо от
 *   того, менялись данные или нет (старое значение с новым не совпадёт по построению). Это не
 *   сбой; заодно этим прогоном доедут потерянные правки. Дальше всё как обычно.
 * ⚠ Цена: JSON.stringify теперь считается на КАЖДОМ прогоне, а не только при публикации
 *   (раньше он стоял после проверки). Времена пишутся в журнал каждым прогоном - см. «чтение /
 *   сумма» в строке лога; при 96 прогонах в сутки следить, чтобы суточный расход не подошёл
 *   к квоте 90 минут.
 *
 * ОТКАТ (полный, за минуту)
 *   Снять триггер publishSnapshot и удалить ветку snapshots в консоли Firebase. Всё.
 *   Дашборд этого не заметит - он на снимок ещё не смотрит, а мастер и лист 00.UnifiedData
 *   скрипт не трогает вовсе: читает и не пишет.
 *
 * ─────────────────────────────────────────────────────────────────────────────────────────────
 * РАСКЛАДКА В БАЗЕ и ПОЧЕМУ ОНА ТАКАЯ
 *   snapshots/rawdata/meta            = {gen, rows, cols, chunks, bytes, updatedAt, srcRows}
 *   snapshots/rawdata/gen/<gen>/0..N  = куски base64 по ~256 КБ
 *
 * ⚠ Бесплатный тариф RTDB даёт 10 ГБ исходящего в месяц. Полная перекачка снимка на каждое
 *   обновление при трёх пользователях онлайн и сотне обновлений в день - это больше 30 ГБ,
 *   в тариф не влезает. Поэтому в заходе 2 клиент качает поколение ОДИН раз и дальше живёт
 *   на дельтах. Здесь дельт ещё нет, но раскладка заложена под них: meta отдельно от данных
 *   (клиент слушает крошечный указатель, а не 4 МБ), данные кусками, поколение помечено.
 * ⚠ ПОРЯДОК ПУБЛИКАЦИИ ОБЯЗАТЕЛЕН: СНАЧАЛА куски, ПОТОМ meta. meta - это указатель, клиент по
 *   нему пойдёт за данными; обнови её первой - и он придёт за кусками, которых ещё нет.
 * ⚠ Ключи кусков числовые (0..N), и RTDB отдаст их клиенту МАССИВОМ, а не объектом (та же
 *   ловушка, что с config/plansBaseline). Читателю в заходе 2 быть толерантным к обеим формам.
 *   Здесь поколение пишется ОДНИМ PUT: так оно появляется целиком или не появляется вовсе -
 *   полуписаного поколения, на которое уже показывает meta, быть не может.
 *   Храним текущее и предыдущее поколение (KEEP_GENS), остальные удаляем - иначе база распухнет.
 *
 * ФОРМАТ СНИМКА - массив массивов, ровно тот, что сейчас приходит из Sheets API в поле values,
 * ВКЛЮЧАЯ строку заголовков. Никаких преобразований в объекты: дашборд уже умеет работать
 * с этим видом, и любое «улучшение» формата означало бы правку расчётного слоя.
 * ⚠ ДАТЫ. Sheets API дашборд зовёт с dateTimeRenderOption=SERIAL_NUMBER и parseDate ждёт именно
 *   серийное число (index.html:1277), а getValues() отдаёт объекты Date, которые JSON.stringify
 *   превратил бы в строку ISO. Поэтому Date здесь переводится в серийный номер РУЧНОГО
 *   разбора (serial_) - иначе снимок «те же данные» отличался бы от Sheets API в двух колонках
 *   дат, а это половина отчётов.
 *
 * @see freeze-rawdata.gs - заливка RawData значениями, другой скрипт и другая задача; общего
 *      состояния у них нет, пересечение по времени разводится LockService.
 */

// ── Источник ────────────────────────────────────────────────────────────────────────────────
var MASTER_ID  = '1Jyv6nE-Gy-Fwz-sEWWa3DvzbfYEGBa6sRryuvmI8vJo';
var SRC_SHEET  = '00.UnifiedData';
var COLS       = 15;   // ровно A:O, порядок как в COL (index.html)
var C_INCOME   = 4;    // COL.INCOME  - колонка E, для контрольной суммы
var C_OUTCOME  = 5;    // COL.OUTCOME - колонка F

// ── Приёмник ────────────────────────────────────────────────────────────────────────────────
var DB_URL     = 'https://upr-finance-opora-8521b-default-rtdb.europe-west1.firebasedatabase.app';
var ROOT       = 'snapshots/rawdata';
var CHUNK      = 256 * 1024;  // символов base64 в куске
var KEEP_GENS  = 2;           // текущее + предыдущее

// ── Сторож ──────────────────────────────────────────────────────────────────────────────────
var MIN_ROWS    = 1000;   // абсолютный минимум
var ROWS_RATIO  = 0.98;   // не публиковать, если строк стало меньше 98% от прошлой публикации
var BAD_MARKERS = ['загрузка', 'loading', '#n/a', '#error!', '#ref!', '#value!', '#name?'];
var ALERT_TO    = 'o017ev@gmail.com';
var ALERT_EVERY = 60 * 60 * 1000; // не чаще раза в час

var EPOCH_DAYS = 25569;   // дней между 1899-12-30 (ноль серийной шкалы) и 1970-01-01

/** Основная: публикует снимок, если данные изменились и сторож не возразил. Её и вешать на триггер. */
function publishSnapshot() { return run_(false); }

/**
 * Без сторожа и без проверки «данные не менялись» - публикует что есть.
 * ⚠ ТОЛЬКО РУКАМИ и только осознанно: на триггер не вешать. Нужна ровно в двух случаях -
 *   первый прогон на пустой базе и восстановление после того, как ветку snapshots удалили.
 */
function publishSnapshotForce() { return run_(true); }

/** Диагностика: печатает в журнал meta из базы и локальную контрольную сумму. Ничего не пишет. */
function snapshotStatus() {
  var meta = rtdbGet_(ROOT + '/meta');
  Logger.log('meta в базе: ' + (meta ? JSON.stringify(meta) : '(нет - снимок ещё не публиковался)'));

  var p = PropertiesService.getScriptProperties();
  Logger.log('локально запомнено: сумма=' + (p.getProperty('sum') || '-') +
             ' · поколение=' + (p.getProperty('gen') || '-') +
             ' · публикация=' + (p.getProperty('publishedAt') || '-'));
  var lastAbort = p.getProperty('lastAbort');
  if (lastAbort) Logger.log('последний отказ: ' + lastAbort);

  var t0  = Date.now();
  var src = read_();
  var tRead = Date.now() - t0;
  var t1  = Date.now();
  var sum = sum_(src.rows, JSON.stringify(src.rows));
  var tSum = Date.now() - t1;
  Logger.log('в мастере сейчас: ' + src.rows.length + ' строк (лист - ' + src.srcRows + '), сумма=' + sum);
  Logger.log('время: чтение ' + tRead + ' мс / сумма (stringify+MD5) ' + tSum + ' мс');
  if (meta && meta.rows !== src.rows.length) {
    Logger.log('⚠ расхождение: в снимке ' + meta.rows + ' строк, в мастере ' + src.rows.length);
  }
  return meta;
}

// ─────────────────────────────────────────────────────────────────────────────────────────────

function run_(force) {
  // LockService на ВЕСЬ прогон: не пересечься ни с архивным скриптом, ни с самим собой
  // (15-минутный триггер против прогона, который затянулся на сжатии 23 МБ).
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) { Logger.log('Прогон пропущен: занят другим выполнением'); return null; }

  try {
    var p  = PropertiesService.getScriptProperties();
    var t0 = Date.now();

    var src  = read_();
    var rows = src.rows;
    var tRead = Date.now() - t0;

    // ⚠ JSON считается ЗДЕСЬ, ДО проверки «не менялось»: контрольная сумма - MD5 от него (см. шапку).
    //   Значит его цену платит каждый прогон, а не только публикующий; отсюда и замеры в журнале.
    //   Gzip остаётся ниже - его до проверки считать незачем, он дороже всего в скрипте.
    var t1   = Date.now();
    var json = JSON.stringify(rows);
    var sum  = sum_(rows, json);
    var tSum = Date.now() - t1;
    var times = 'чтение ' + tRead + ' мс / сумма ' + tSum + ' мс';

    // meta прошлой публикации нужна и сторожу (сравнение числа строк), и проверке «не менялось»
    var prev = rtdbGet_(ROOT + '/meta');

    // ── Данные не менялись - выходим, не трогая базу. На этот случай придётся БОЛЬШИНСТВО прогонов.
    // ⚠ Сверяемся не только с локальной суммой, но и с тем, что снимок реально ЛЕЖИТ в базе и это
    //   наше поколение: иначе после удаления ветки (или чужой перезаписи) сумма совпадала бы вечно,
    //   и скрипт молча не публиковал бы никогда - самый неприятный из возможных отказов, немой.
    if (!force && sum === p.getProperty('sum') && prev && prev.gen === p.getProperty('gen')) {
      Logger.log('Данные не менялись (' + rows.length + ' строк, сумма ' + sum + ') - публикации нет · ' +
                 times + ' · прогон ' + (Date.now() - t0) + ' мс');
      return null;
    }

    if (!force) {
      var why = guard_(rows, prev, p);
      if (why) return abort_(why, rows, prev);
    }

    // ── Упаковка. gzip → base64. Тот же массив, что отдаёт Sheets API в values.
    var b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(json, 'application/json')).getBytes());
    var parts = {};
    var n = 0;
    for (var i = 0; i < b64.length; i += CHUNK) { parts[String(n++)] = b64.substr(i, CHUNK); }

    var tz  = src.tz;
    var now = new Date();
    var gen = Utilities.formatDate(now, tz, 'yyyyMMdd-HHmmss');

    // ── Публикация: СНАЧАЛА данные, ПОТОМ указатель (см. шапку - порядок обязателен).
    rtdbPut_(ROOT + '/gen/' + gen, parts);
    var meta = {
      gen: gen,
      rows: rows.length,           // строк в снимке, включая заголовок
      cols: COLS,
      chunks: n,
      bytes: b64.length,           // размер base64 - столько скачает клиент
      updatedAt: now.toISOString(),
      srcRows: src.srcRows         // строк на листе-источнике (до отсечения пустого хвоста)
    };
    rtdbPut_(ROOT + '/meta', meta);

    // Локальные отметки - ТОЛЬКО после успешной публикации: иначе следующий прогон счёл бы
    // непрошедший снимок опубликованным и не повторил бы попытку (тот же инвариант, что
    // у зеркала localStorage в дашборде - хранится ПОСЛЕДНЕЕ ИЗВЕСТНОЕ ОБЛАКО).
    p.setProperties({
      sum: sum,
      gen: gen,
      head: headKey_(rows[0]),
      publishedAt: now.toISOString()
    }, false);

    prune_(gen, prev && prev.gen);
    Logger.log('Снимок опубликован: поколение ' + gen + ' · ' + rows.length + ' строк · ' +
               n + ' кусков · ' + Math.round(b64.length / 1024) + ' КБ (json ' +
               Math.round(json.length / 1024 / 1024) + ' МБ) · ' + times +
               ' · прогон ' + (Date.now() - t0) + ' мс');
    return meta;

  } catch (e) {
    // Сюда попадают отказы RTDB и сбои чтения мастера. Молчать нельзя: триггер живёт сам по себе,
    // и в журнал никто не смотрит, пока что-нибудь не сломается на стороне дашборда.
    Logger.log('ОШИБКА публикации: ' + (e && e.message ? e.message : e));
    alert_('Снимок не опубликован: ошибка',
           'Скрипт публикации снимка упал с ошибкой:\n\n' + (e && e.stack ? e.stack : e) +
           '\n\nПрежний снимок в базе не тронут - дашборд, когда научится его читать, ' +
           'продолжит работать на нём.');
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Чтение мастера напрямую. Возвращает {rows, srcRows, tz}.
 * ⚠ Диапазон берётся ЖЁСТКО A:O по позициям, а не getDataRange(): лишняя служебная колонка,
 *   дописанная справа в мастере, не должна доехать до дашборда - он читает по индексам.
 */
function read_() {
  var ss = SpreadsheetApp.openById(MASTER_ID);
  var sh = ss.getSheetByName(SRC_SHEET);
  if (!sh) throw new Error('В мастере нет листа «' + SRC_SHEET + '»');

  var tz = ss.getSpreadsheetTimeZone() || 'Europe/Moscow';
  // ⚠ Серийные номера дат считаются по компонентам локального времени скрипта. Разъедься
  //   таймзона скрипта с таймзоной книги - все даты уехали бы на разницу поясов МОЛЧА.
  var stz = Session.getScriptTimeZone();
  if (stz && stz !== tz) {
    throw new Error('Таймзона скрипта (' + stz + ') не совпадает с таймзоной книги (' + tz +
                    '). Даты в снимке уехали бы на разницу поясов. Выровнять в настройках проекта.');
  }

  var last = sh.getLastRow();
  if (!last) throw new Error('Лист «' + SRC_SHEET + '» пуст');
  var vals = sh.getRange(1, 1, last, COLS).getValues();

  // Пустой хвост отсекаем: подтянутые вниз формулы в мастере оставляют строки из пятнадцати ''.
  var end = vals.length;
  while (end > 0 && isBlankRow_(vals[end - 1])) end--;
  vals.length = end;

  for (var r = 0; r < vals.length; r++) {
    var row = vals[r];
    for (var c = 0; c < COLS; c++) {
      var v = row[c];
      if (v instanceof Date) row[c] = serial_(v);
      else if (v === null || v === undefined) row[c] = '';
    }
  }
  return { rows: vals, srcRows: last, tz: tz };
}

/** Дата → серийный номер Sheets (то же, что отдаёт API при dateTimeRenderOption=SERIAL_NUMBER). */
function serial_(d) {
  var days = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000 + EPOCH_DAYS;
  var secs = d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds();
  return secs ? days + secs / 86400 : days;
}

function isBlankRow_(row) {
  for (var c = 0; c < COLS; c++) { if (row[c] !== '' && row[c] !== null && row[c] !== undefined) return false; }
  return true;
}

/**
 * Контрольная сумма: строк | Σ поступлений | Σ списаний | MD5 всего массива.
 * ⚠ РЕШАЕТ сумма ТОЛЬКО последняя часть - MD5 от того самого JSON, который уезжает в базу.
 *   Первые три - для человека в журнале: MD5 не говорит ничего, а по ним сразу видно масштаб
 *   изменения. До 16.08.2026 суммой были только они, и правка любой нечисловой ячейки скрипт
 *   не будила - см. подробности в шапке файла.
 * Копейки в суммах округляем: на 60 тыс. слагаемых хвост float гуляет в последних разрядах.
 * ⚠ json приходит АРГУМЕНТОМ, а не считается здесь: он же нужен вызывающему для gzip, и второй
 *   stringify на 23,5 МБ - это лишние секунды на каждом прогоне.
 */
function sum_(rows, json) {
  var inc = 0, out = 0;
  for (var r = 1; r < rows.length; r++) {
    inc += Number(rows[r][C_INCOME])  || 0;
    out += Number(rows[r][C_OUTCOME]) || 0;
  }
  var md5 = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, json)
    .map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); })
    .join('');
  return rows.length + '|' + inc.toFixed(2) + '|' + out.toFixed(2) + '|' + md5;
}

/** Нормализованная строка заголовков - тем же способом, что сверяет дашборд (NBSP, регистр, пробелы). */
function headKey_(head) {
  var a = [];
  for (var c = 0; c < COLS; c++) {
    a.push(String(head && head[c] != null ? head[c] : '')
      .replace(/ /g, ' ').replace(/\s+/g, ' ').trim().toLowerCase());
  }
  return a.join('|');
}

/**
 * Сторож. Возвращает причину отказа или null.
 * ⚠ Здесь он критичнее, чем в архивном скрипте: битый архив просто лежит, а битый снимок
 *   за полминуты разъедется по всем открытым дашбордам. Поэтому «не опубликовать» - всегда
 *   лучше, чем «опубликовать сомнительное»: прежнее поколение остаётся на месте и живо.
 */
function guard_(rows, prev, p) {
  if (rows.length < MIN_ROWS) {
    return 'строк всего ' + rows.length + ' при минимуме ' + MIN_ROWS;
  }
  if (prev && prev.rows && rows.length < prev.rows * ROWS_RATIO) {
    return 'строк стало ' + rows.length + ' против ' + prev.rows + ' в прошлой публикации (порог - ' +
           Math.round(ROWS_RATIO * 100) + '%)';
  }

  var head = headKey_(rows[0]);
  var was  = p.getProperty('head');
  if (was && was !== head) {
    var a = was.split('|'), b = head.split('|'), diff = [];
    for (var i = 0; i < COLS; i++) { if (a[i] !== b[i]) diff.push('#' + (i + 1) + ': «' + (a[i] || '-') + '» → «' + (b[i] || '-') + '»'); }
    return 'изменилась строка заголовков (' + diff.join('; ') + ')';
  }
  // Шапка целиком пустой/битой не бывает у живого листа - это признак того, что мастер
  // читается в момент пересборки. Первый прогон эталона ещё не имеет, и без этой проверки
  // он бы такое состояние ЗАПОМНИЛ как эталон.
  var hc = head.split('|');
  for (var k = 0; k < COLS; k++) {
    if (hc[k] === '' || hc[k].charAt(0) === '#') return 'шапка невалидна (колонка #' + (k + 1) + ': «' + hc[k] + '»)';
  }

  // Маркеры незавершённого импорта в первой колонке (там дата ДДС - текст в ней уже подозрителен).
  for (var r = 1; r < rows.length; r++) {
    var v = rows[r][0];
    if (typeof v !== 'string' || !v) continue;
    var s = v.toLowerCase();
    for (var m = 0; m < BAD_MARKERS.length; m++) {
      if (s.indexOf(BAD_MARKERS[m]) >= 0) {
        return 'в первой колонке строки ' + (r + 1) + ' лежит «' + v + '» - источник ещё считается';
      }
    }
  }
  return null;
}

/** Отказ: прежний снимок ОСТАЁТСЯ в базе. Плюс письмо, но не чаще раза в час. */
function abort_(why, rows, prev) {
  Logger.log('ОТКАЗ от публикации: ' + why);
  PropertiesService.getScriptProperties().setProperty('lastAbort', new Date().toISOString() + ' · ' + why);
  alert_('Снимок данных НЕ опубликован',
    'Скрипт публикации снимка отказался писать в базу.\n\n' +
    'Причина: ' + why + '.\n\n' +
    'Было в прошлой публикации: ' + (prev && prev.rows ? prev.rows + ' строк (' + prev.updatedAt + ')' : 'публикаций ещё не было') + '.\n' +
    'Стало сейчас в мастере: ' + rows.length + ' строк.\n\n' +
    'ПРЕЖНИЙ СНИМОК ЦЕЛ и лежит в базе - ничего не потеряно. Скрипт повторит попытку на следующем ' +
    'прогоне (раз в 15 минут) и опубликует сам, как только данные придут в порядок. ' +
    'Если это ожидаемое изменение (например, колонки в мастере переставили намеренно) - ' +
    'запустить publishSnapshotForce() руками из редактора скриптов.');
  return null;
}

/** Письмо с ограничением частоты: 15-минутный триггер иначе завалил бы почту за ночь. */
function alert_(subject, body) {
  try {
    var p = PropertiesService.getScriptProperties();
    var last = Number(p.getProperty('lastAlertAt') || 0);
    if (Date.now() - last < ALERT_EVERY) { Logger.log('Письмо подавлено (было менее часа назад)'); return; }
    MailApp.sendEmail({ to: ALERT_TO, subject: '[Дашборд] ' + subject, body: body });
    p.setProperty('lastAlertAt', String(Date.now()));
  } catch (e) {
    Logger.log('Письмо отправить не удалось: ' + (e && e.message ? e.message : e));
  }
}

/** Чистка старых поколений. Держим текущее и предыдущее - на предыдущем ещё могут сидеть клиенты. */
function prune_(cur, prevGen) {
  var list = rtdbGet_(ROOT + '/gen', 'shallow=true'); // ⚠ именно shallow: иначе выкачали бы все поколения
  if (!list) return;
  var keep = {};
  keep[cur] = 1;
  if (prevGen) keep[prevGen] = 1;
  var names = Object.keys(list).sort();
  // Если поколений накопилось больше KEEP_GENS (после ручных прогонов) - оставляем самые свежие.
  while (names.length && Object.keys(keep).length < KEEP_GENS) { keep[names[names.length - 1]] = 1; names.pop(); }
  for (var i = 0; i < names.length; i++) {
    if (keep[names[i]]) continue;
    rtdbDel_(ROOT + '/gen/' + names[i]);
    Logger.log('Удалено старое поколение: ' + names[i]);
  }
}

// ── RTDB REST. Авторизация - токеном самого скрипта, служебного аккаунта и ключей нет ────────
function rtdbGet_(path, qs) {
  var r = fetch_(path + '.json' + (qs ? '?' + qs : ''), { method: 'get' });
  var t = r.getContentText();
  return (!t || t === 'null') ? null : JSON.parse(t);
}
function rtdbPut_(path, val) {
  fetch_(path + '.json', { method: 'put', contentType: 'application/json', payload: JSON.stringify(val) });
}
function rtdbDel_(path) {
  fetch_(path + '.json', { method: 'delete' });
}
function fetch_(path, opt) {
  opt.headers = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  opt.muteHttpExceptions = true; // код ответа разбираем сами - иначе в журнал уедет HTML-страница ошибки
  var r = UrlFetchApp.fetch(DB_URL + '/' + path, opt);
  var c = r.getResponseCode();
  if (c < 200 || c >= 300) {
    throw new Error('RTDB ' + opt.method.toUpperCase() + ' ' + path + ' → HTTP ' + c + ': ' +
                    String(r.getContentText()).slice(0, 400));
  }
  return r;
}
