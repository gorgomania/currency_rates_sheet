const CURRENCIES = ['EUR', 'GBP', 'CHF', 'CAD', 'AUD'];
const THRESHOLD = 0.02; // 2%
const API = 'https://api.frankfurter.dev/v1/';

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName('data') || ss.insertSheet('data');
}

// Один раз вручную: очищает лист и грузит историю
function backfill() {
  const sheet = getSheet_();
  sheet.clear();
  sheet.appendRow(['time', 'currency', 'rate']);
  const url = API + '2026-09-01..?base=USD&symbols=' + CURRENCIES.join(',');
  const data = JSON.parse(UrlFetchApp.fetch(url).getContentText());
  const rows = [];
  for (const [date, rates] of Object.entries(data.rates)) {
    for (const [cur, rate] of Object.entries(rates)) rows.push([new Date(date), cur, rate]);
  }
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  const last = Object.values(data.rates).pop();
  PropertiesService.getScriptProperties().setProperty('lastRates', JSON.stringify(last));
}

// По триггеру ежедневно в 9:00
function updateRates() {
  const url = API + 'latest?base=USD&symbols=' + CURRENCIES.join(',');
  const data = JSON.parse(UrlFetchApp.fetch(url).getContentText());
  const sheet = getSheet_();
  const date = new Date(data.date);

  // ЕЦБ не опубликовал новых данных (выходные, праздники) - выходим
  const lastDate = sheet.getRange(sheet.getLastRow(), 1).getValue();
  if (lastDate instanceof Date && lastDate.getTime() === date.getTime()) return;

  const rows = Object.entries(data.rates).map(([cur, rate]) => [date, cur, rate]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
  sendDigest_(data.date, data.rates);
}

function sendDigest_(dateStr, rates) {
  const props = PropertiesService.getScriptProperties();
  const prev = JSON.parse(props.getProperty('lastRates') || '{}');

  const lines = Object.entries(rates).map(([cur, rate]) => {
    if (!prev[cur]) return `${cur}: ${rate}`;
    const change = (rate - prev[cur]) / prev[cur];
    const arrow = change > 0 ? '▲' : change < 0 ? '▼' : '■';
    const sign = change > 0 ? '+' : '';
    const flag = Math.abs(change) >= THRESHOLD ? ' ⚠️' : '';
    return `${cur}: ${rate} ${arrow} ${sign}${(change * 100).toFixed(2)}%${flag}`;
  });

  props.setProperty('lastRates', JSON.stringify(rates));
  sendTelegram_(`Курсы за ${dateStr} (единиц валюты за 1 USD, изменение к предыдущему дню):\n` + lines.join('\n'));
}

function testDigest() {
  const props = PropertiesService.getScriptProperties();
  const backup = props.getProperty('lastRates');

  props.setProperty('lastRates', JSON.stringify({ EUR: 0.80, GBP: 0.75, CHF: 0.83 }));
  sendDigest_('тест', { EUR: 0.89, GBP: 0.76, CHF: 0.83 });

  if (backup) props.setProperty('lastRates', backup);
  else props.deleteProperty('lastRates');
}

function sendTelegram_(text) {
  const p = PropertiesService.getScriptProperties();
  const resp = UrlFetchApp.fetch(`https://api.telegram.org/bot${p.getProperty('TG_TOKEN')}/sendMessage`, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: p.getProperty('TG_CHAT_ID'), text }),
    muteHttpExceptions: true,
  });
  Logger.log(resp.getContentText());
}

function testTelegram() { sendTelegram_('Тест: бот работает'); }