import { DateTime } from 'luxon';

type UpdateData = {
  mailDate: DateTime,
  targetDate: DateTime,
  quantity?: number,
  price: {
    average?: number,
    s4?: number, s5?: number, sl?: number, sm?: number,
    y4?: number, y5?: number, yl?: number, ym?: number
  }
};

const gas: any = global;

gas._main = () => {
  const WATERMELON_SPREAD_SHEET_ID = getProperty('WATERMELON_SPREAD_SHEET_ID');

  // メールを検索する条件
  const SEARCH_KEYWORD = 'label:市況-スイカ';
  // 設定シートのメール検索日のセル
  const SETTINGS_SHEET_SEARCH_MAIL_DATE = 'B2';
  // Webhook Url
  const WEBHOOK_URLS = getProperty('WEBHOOK_URLS').split('|');

  const spreadSheet = SpreadsheetApp.openById(WATERMELON_SPREAD_SHEET_ID);
  const settingsSheet = spreadSheet.getSheetByName('SETTINGS');
  if (!settingsSheet) throw new Error('SETTINGSシートが存在しません');

  const searchMailDateRange = settingsSheet.getRange(SETTINGS_SHEET_SEARCH_MAIL_DATE);
  const searchMailDateValue: string = searchMailDateRange.getValue();

  let searchMailDate: DateTime | undefined = undefined;
  let latestMailDate: DateTime | undefined = undefined;

  // メールの検索キーワードを組み立て
  let searchKeyword = SEARCH_KEYWORD;
  if (searchMailDateValue) {
    searchMailDate = DateTime.fromISO(searchMailDateValue);
    // 最終検索日以降
    searchKeyword += ' after:' + searchMailDate.toFormat('yyyy/MM/dd');
  }

  let messages: GoogleAppsScript.Gmail.GmailMessage[] = [];
  for (const thread of GmailApp.search(searchKeyword)) {
    for (const message of thread.getMessages()) {
      messages.push(message);
    }
  }

  messages = messages.sort((a, b) => a.getDate().getTime() - b.getDate().getTime());

  Logger.log('searchKeyword: %s, messageCount: %s', searchKeyword, messages.length);

  let updateDatas: UpdateData[] = [];

  // メールから市況データを集計
  for (const message of messages) {
    const plainBody = message.getPlainBody();
    const nextLineGenerator = (function* () {
      for (let line of plainBody.split('\r\n')) {
        line = normalize(line.trim());
        if (line) yield line;
      }
    })();
    const readBody = () => {
      const value = nextLineGenerator.next().value;
      return value ? value : '';
    };

    // Mail
    const mailDate = DateTime.fromMillis(message.getDate().getTime());
    if (searchMailDate && mailDate <= searchMailDate) continue;
    latestMailDate = mailDate;

    for (const url of WEBHOOK_URLS) {
      try {
        UrlFetchApp.fetch(url, {
          method: 'post',
          payload: {
            username: message.getSubject(),
            content: normalize(plainBody)
          }
        });
      } catch (error) {
        console.error(error);
      }
    }

    const mailMonth = mailDate.month;
    // mm月dd日出荷
    const linePd = readBody();
    const pdMatcher = linePd.match(/(.+)月\s*(.+)日出荷/);
    if (!pdMatcher) return;
    // AL, AM, AS
    const lineS4 = readBody();
    const lineS5 = readBody();
    const lineSl = readBody();
    const lineSm = readBody();
    const lineY4 = readBody();
    const lineY5 = readBody();
    const lineYl = readBody();
    const lineYm = readBody();
    // label 平均単価
    readBody();
    // n円
    const lineAvg = readBody();
    // label 出荷箱数
    readBody();
    // n箱
    const lineQty = readBody();
    // 本文に年がないので、メールの時刻から取得する
    let year = mailDate.year;
    const month = Number(pdMatcher[1]);
    const day = Number(pdMatcher[2]);
    // 前年の市況が年初に送られてきた場合
    if (month === 12 && mailMonth === 1) year--;

    const updateData: UpdateData = {
      mailDate: mailDate,
      targetDate: DateTime.local(year, month, day),
      quantity: formatNumber(lineQty),
      price: {
        average: formatNumber(lineAvg),
        s4: formatNumber(lineS4), s5: formatNumber(lineS5),
        sl: formatNumber(lineSl), sm: formatNumber(lineSm),
        y4: formatNumber(lineY4), y5: formatNumber(lineY5),
        yl: formatNumber(lineYl), ym: formatNumber(lineYm)
      }
    };

    updateDatas.push(updateData);
  };

  updateDatas = updateDatas.sort((a, b) => a.targetDate.diff(b.targetDate).milliseconds);

  // シートに書き出し
  for (const updateData of updateDatas) {
    const sheetName = String(updateData.targetDate.year);
    let sheet = spreadSheet.getSheetByName(sheetName);

    // シートが存在しない場合、雛形からコピーして作成する
    if (!sheet) {
      const templateSheet = spreadSheet.getSheetByName('TEMPLATE');
      if (!templateSheet) throw new Error('SETTINGSシートが存在しません');

      sheet = templateSheet.copyTo(spreadSheet);
      spreadSheet.setActiveSheet(sheet);
      spreadSheet.moveActiveSheet(1);
      sheet.setName(sheetName).showSheet();
    }

    const row = sheet.getLastRow() + 1;
    let column = 1;
    sheet.getRange(row, column++).setValue(updateData.targetDate.toFormat('yyyy/MM/dd'));
    sheet.getRange(row, column++).setValue(updateData.price.s4);
    sheet.getRange(row, column++).setValue(updateData.price.s5);
    sheet.getRange(row, column++).setValue(updateData.price.sl);
    sheet.getRange(row, column++).setValue(updateData.price.sm);
    sheet.getRange(row, column++).setValue(updateData.price.y4);
    sheet.getRange(row, column++).setValue(updateData.price.y5);
    sheet.getRange(row, column++).setValue(updateData.price.yl);
    sheet.getRange(row, column++).setValue(updateData.price.ym);
    sheet.getRange(row, column++).setValue(updateData.price.average);
    sheet.getRange(row, column++).setValue(updateData.quantity);
    sheet.getRange(row, column++).setValue(updateData.mailDate.toFormat('yyyy/MM/dd HH:mm:ss'));
  };

  // 全てが正常終了したら、設定シートを更新する
  if (latestMailDate) {
    searchMailDateRange.setValue(latestMailDate.toISO());
  }
};

function getProperty(key: string, defaultValue?: any): string {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (value) return value;
  if (defaultValue) return defaultValue;
  throw new Error(`Undefined property: ${key}`);
}

function normalize(s: string) {
  // F*ck Zenkaku
  return s.replace(/[Ａ-Ｚａ-ｚ０-９]/g,
    s => String.fromCharCode(s.charCodeAt(0) - 65248)).replace(/　/g, ' ');
}

function formatNumber(s: string): number | undefined {
  const sn = s.replace(/\D*(\d*)\D*/, '$1');
  if (!sn) return undefined;
  return Number(sn);
}
