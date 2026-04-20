const ss = SpreadsheetApp.openById('YOUR_SPREADSHEET_ID'); // ← スプレッドシートIDを設定してください

/**
 * メニュー追加
 * シート選択用サイドバーを表示するメニューを追加
 */
function addMenu() {
  SpreadsheetApp.getUi()
  .createMenu('稼働管理システム',)
  .addItem('サイドバー表示','showSidebar')
  .addToUi();
}

/**
 * サイドバー表示
 */
function showSidebar() {
  let sidebar = HtmlService.createTemplateFromFile('index');
  sidebar.data = getSheets();
  sidebar = sidebar.evaluate();
  SpreadsheetApp.getUi().showSidebar(sidebar);
}

/**
 * シート名のリストを取得
 * @return {array} シート名の配列
 */
function getSheets() {
  const data = ss.getSheets().map(value => value.getSheetName());
  return data;
}

/**
 * 稼働者リストを取得
 * @param {string} sheetName　シート名
 * @return {array} 該当シートのA4:Pの値の二次元配列
 */
function getList(sheetName) {
  const sh = ss.getSheetByName(sheetName);
  try {
    const lastR = sh.getRange(500,5).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    let data = sh.getRange(4,1,lastR - 3,16).getValues();
    data.forEach((value,index) => {
      if (Object.prototype.toString.call(value[8]) == '[object Date]') {
        value[8] = Utilities.formatDate(value[8],"JST","M月d日");
      }
      if (Object.prototype.toString.call(value[9]) == '[object Date]') {
        value[9] = Utilities.formatDate(value[9],"JST","M月d日");
      }
      value[16] = index;
    });
    return JSON.stringify([data,sheetName]);
  } catch(e) {
    console.log(e);
    return JSON.stringify([[''],""]);
  }
}

/**
 * メールを送信する
 * @param {string} month 月
 * @param {array} data 稼働者の配列
 * @param {string} sheetName シート名
 * @return {}
 */
function sendMail(month, data, sheetName) {
  for (let r of data) {
    const to = r[6];
    const name = r[5];
    const report = r[8];
    const invoice = r[9];

    const subject = `【${name}様 ${month}月分：作業報告書・請求書の提出依頼】YOUR_COMPANY_NAMEです`; // ← 会社名を設定してください
    const text = `${name}様

      いつもお世話になっております。
      YOUR_COMPANY_NAMEの YOUR_NAME でございます。

      ${month}月も業務支援いただき誠にありがとうございます。
      表題の件につきまして、下記スケジュールでSlackもしくは、本メール全返信にてご提出いただければ幸いです。
      -----------------------------------
      ・作業報告書：${report}まで
      ・請求書　　：${invoice}まで
      -----------------------------------
      ご不明点やご相談等ございましたらお気軽にお申し付けくださいませ。

      お手数をおかけしますが、ご確認の程お願いいたします。
      何卒よろしくお願いいたします。

      -----------------------------------
      YOUR_COMPANY_NAME　YOUR_NAME
      Mail: your-email@example.com       // ← 連絡先メールアドレスを設定してください
      TEL: 000-0000-0000                 // ← 電話番号を設定してください
      HP : https://your-company.example.com  // ← 会社HPを設定してください
      〒000-0000　YOUR_ADDRESS           // ← 住所を設定してください
      `.replace(/^\s+$|^ {6}/gm,'');

    const option = {
      cc: 'cc1@example.com,cc2@example.com', // ← CCアドレスを設定してください
      name: 'YOUR_COMPANY_NAME　YOUR_NAME'   // ← 送信者名を設定してください
    };
    const draft = GmailApp.createDraft(to, subject, text, option);
    // draft.send();
    ss.getSheetByName(sheetName).getRange(r[16]+4, 4).setValue('済');
  }
  return;
}
