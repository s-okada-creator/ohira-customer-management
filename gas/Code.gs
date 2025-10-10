/**
 * 顧客データ管理システム - Google Apps Script
 * 新規顧客追加機能
 */

/**
 * スプレッドシートを開いた時にメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('顧客管理')
    .addItem('新規顧客を追加', 'showAddCustomerDialog')
    .addItem('顧客を検索', 'showSearchCustomerDialog')
    .addItem('統計情報を表示', 'showStatistics')
    .addToUi();
}

/**
 * 新規顧客追加ダイアログを表示
 */
function showAddCustomerDialog() {
  const html = HtmlService.createHtmlOutputFromFile('AddCustomerDialog')
    .setWidth(600)
    .setHeight(500)
    .setTitle('新規顧客追加');
  
  SpreadsheetApp.getUi().showModalDialog(html, '新規顧客を追加');
}

/**
 * 新規顧客を追加するメイン処理
 */
function addNewCustomer(customerData) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 新しい顧客番号を生成
    const newCustomerNumber = generateNewCustomerNumber();
    
    // 顧客名簿シートを作成
    const customerNameSheet = createCustomerNameSheet(spreadsheet, newCustomerNumber, customerData);
    
    // 顧客データベースシートを作成
    const customerDatabaseSheet = createCustomerDatabaseSheet(spreadsheet, newCustomerNumber, customerData);
    
    // 目次シートを更新
    updateIndexSheet(spreadsheet, newCustomerNumber, customerData, [customerNameSheet.getName(), customerDatabaseSheet.getName()]);
    
    // 処理ログを更新
    updateProcessLog(spreadsheet, '新規顧客追加', newCustomerNumber);
    
    // 新しく作成されたシートに移動
    spreadsheet.setActiveSheet(customerNameSheet);
    
    SpreadsheetApp.getUi().alert('新規顧客が正常に追加されました！\n顧客番号: ' + newCustomerNumber);
    
    return {
      success: true,
      customerNumber: newCustomerNumber,
      message: '新規顧客が正常に追加されました'
    };
    
  } catch (error) {
    console.error('新規顧客追加エラー:', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
    return {
      success: false,
      message: 'エラーが発生しました: ' + error.message
    };
  }
}

/**
 * 新しい顧客番号を生成
 */
function generateNewCustomerNumber() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = spreadsheet.getSheetByName('目次');
  
  if (!indexSheet) {
    return '0001';
  }
  
  const lastRow = indexSheet.getLastRow();
  let maxNumber = 0;
  
  // 既存の顧客番号から最大値を取得
  for (let i = 2; i <= lastRow; i++) {
    const customerNumber = indexSheet.getRange(i, 1).getValue();
    if (customerNumber && typeof customerNumber === 'string') {
      const number = parseInt(customerNumber);
      if (!isNaN(number) && number > maxNumber) {
        maxNumber = number;
      }
    }
  }
  
  // 次の番号を生成（4桁ゼロパディング）
  const nextNumber = maxNumber + 1;
  return nextNumber.toString().padStart(4, '0');
}

/**
 * 顧客名簿シートを作成
 */
function createCustomerNameSheet(spreadsheet, customerNumber, customerData) {
  const sheetName = `${customerNumber}_${customerData.customerName}_${customerData.carName}_顧客名簿`;
  const sheet = spreadsheet.insertSheet(sheetName);
  
  // ヘッダー行を設定
  const headers = [
    ['顧客名簿　　　　　　　　　 大平自動車商会', '', '', ''],
    ['コード№', customerNumber, '会社住所2', ''],
    ['名　前', customerData.customerName, '会社電話番号', ''],
    ['ふりがな', customerData.customerNameKana, '車　名', customerData.carName],
    ['自宅郵便番号', customerData.postalCode, '車両ナンバー', customerData.carNumber],
    ['自宅住所1', customerData.address1, '車両年式', customerData.carYear],
    ['自宅住所2', customerData.address2, '車両色', customerData.carColor],
    ['自宅電話番号', customerData.phoneNumber, '購入日', customerData.purchaseDate],
    ['携帯電話番号', customerData.mobileNumber, '保証期間', customerData.warrantyPeriod],
    ['メールアドレス', customerData.email, '次回点検予定', customerData.nextInspection],
    ['備考', customerData.notes, '', ''],
    ['', '', '', '']
  ];
  
  // データを設定
  for (let i = 0; i < headers.length; i++) {
    sheet.getRange(i + 1, 1, 1, 4).setValues([headers[i]]);
  }
  
  // スタイルを設定
  sheet.getRange(1, 1, 1, 4).merge().setFontWeight('bold').setFontSize(14);
  sheet.getRange(2, 1, 1, 4).setFontWeight('bold');
  sheet.getRange(3, 1, 1, 4).setFontWeight('bold');
  sheet.getRange(4, 1, 1, 4).setFontWeight('bold');
  sheet.getRange(5, 1, 1, 4).setFontWeight('bold');
  sheet.getRange(6, 1, 1, 4).setFontWeight('bold');
  sheet.getRange(7, 1, 1, 4).setFontWeight('bold');
  sheet.getRange(8, 1, 1, 4).setFontWeight('bold');
  sheet.getRange(9, 1, 1, 4).setFontWeight('bold');
  sheet.getRange(10, 1, 1, 4).setFontWeight('bold');
  sheet.getRange(11, 1, 1, 4).setFontWeight('bold');
  
  // 列幅を調整
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 200);
  
  return sheet;
}

/**
 * 顧客データベースシートを作成
 */
function createCustomerDatabaseSheet(spreadsheet, customerNumber, customerData) {
  const sheetName = `${customerNumber}_${customerData.customerName}_${customerData.carName}_顧客データベース`;
  const sheet = spreadsheet.insertSheet(sheetName);
  
  // ヘッダー行を設定
  const headers = [
    'コード№', '名　前', 'ふりがな', '自宅郵便番号', '自宅住所1', '自宅住所2',
    '自宅電話番号', '携帯電話番号', 'メールアドレス', '車　名', '車両ナンバー',
    '車両年式', '車両色', '購入日', '保証期間', '次回点検予定', '備考',
    '登録日', '更新日', '担当者', '紹介者', '備考2'
  ];
  
  // データ行を設定
  const dataRow = [
    customerNumber,
    customerData.customerName,
    customerData.customerNameKana,
    customerData.postalCode,
    customerData.address1,
    customerData.address2,
    customerData.phoneNumber,
    customerData.mobileNumber,
    customerData.email,
    customerData.carName,
    customerData.carNumber,
    customerData.carYear,
    customerData.carColor,
    customerData.purchaseDate,
    customerData.warrantyPeriod,
    customerData.nextInspection,
    customerData.notes,
    new Date(),
    new Date(),
    customerData.staff,
    customerData.referrer,
    customerData.notes2
  ];
  
  // ヘッダーを設定
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // データを設定
  sheet.getRange(2, 1, 1, dataRow.length).setValues([dataRow]);
  
  // スタイルを設定
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E0E0E0');
  
  // 列幅を調整
  for (let i = 1; i <= headers.length; i++) {
    sheet.setColumnWidth(i, 100);
  }
  
  return sheet;
}

/**
 * 目次シートを更新
 */
function updateIndexSheet(spreadsheet, customerNumber, customerData, sheetNames) {
  const indexSheet = spreadsheet.getSheetByName('目次');
  
  if (!indexSheet) {
    // 目次シートが存在しない場合は作成
    createIndexSheet(spreadsheet);
    return;
  }
  
  const lastRow = indexSheet.getLastRow();
  
  // 新しい顧客情報を追加
  for (let i = 0; i < sheetNames.length; i++) {
    const sheetName = sheetNames[i];
    const sheetType = sheetName.includes('顧客名簿') ? '顧客名簿' : '顧客データベース';
    
    indexSheet.getRange(lastRow + 1 + i, 1).setValue(customerNumber);
    indexSheet.getRange(lastRow + 1 + i, 2).setValue(customerData.customerName);
    indexSheet.getRange(lastRow + 1 + i, 3).setValue(customerData.carName);
    indexSheet.getRange(lastRow + 1 + i, 4).setValue(sheetType);
    indexSheet.getRange(lastRow + 1 + i, 5).setValue(`新規追加 - ${new Date()}`);
  }
}

/**
 * 目次シートを作成（存在しない場合）
 */
function createIndexSheet(spreadsheet) {
  const indexSheet = spreadsheet.insertSheet('目次', 0);
  
  // ヘッダー行を設定
  const headers = ['顧客番号', '顧客名', '車種名', 'シート名', 'ファイルパス'];
  indexSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // スタイルを設定
  indexSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E0E0E0');
  
  // 列幅を調整
  for (let i = 1; i <= headers.length; i++) {
    indexSheet.setColumnWidth(i, 150);
  }
}

/**
 * 処理ログを更新
 */
function updateProcessLog(spreadsheet, action, customerNumber) {
  let logSheet = spreadsheet.getSheetByName('処理ログ');
  
  if (!logSheet) {
    // 処理ログシートが存在しない場合は作成
    logSheet = spreadsheet.insertSheet('処理ログ');
    
    // ヘッダー行を設定
    const headers = ['処理日時', '処理内容', '顧客番号', '詳細'];
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E0E0E0');
  }
  
  const lastRow = logSheet.getLastRow();
  
  // 新しいログエントリを追加
  logSheet.getRange(lastRow + 1, 1).setValue(new Date());
  logSheet.getRange(lastRow + 1, 2).setValue(action);
  logSheet.getRange(lastRow + 1, 3).setValue(customerNumber);
  logSheet.getRange(lastRow + 1, 4).setValue(`新規顧客追加処理完了`);
}

/**
 * 顧客検索ダイアログを表示
 */
function showSearchCustomerDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SearchCustomerDialog')
    .setWidth(500)
    .setHeight(400)
    .setTitle('顧客検索');
  
  SpreadsheetApp.getUi().showModalDialog(html, '顧客を検索');
}

/**
 * 顧客を検索
 */
function searchCustomer(searchTerm) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = spreadsheet.getSheetByName('目次');
  
  if (!indexSheet) {
    return { success: false, message: '目次シートが見つかりません' };
  }
  
  const lastRow = indexSheet.getLastRow();
  const results = [];
  
  for (let i = 2; i <= lastRow; i++) {
    const customerNumber = indexSheet.getRange(i, 1).getValue();
    const customerName = indexSheet.getRange(i, 2).getValue();
    const carName = indexSheet.getRange(i, 3).getValue();
    const sheetName = indexSheet.getRange(i, 4).getValue();
    
    if (customerNumber && customerName && carName && sheetName) {
      const searchText = `${customerNumber} ${customerName} ${carName}`.toLowerCase();
      if (searchText.includes(searchTerm.toLowerCase())) {
        results.push({
          customerNumber: customerNumber,
          customerName: customerName,
          carName: carName,
          sheetName: sheetName,
          row: i
        });
      }
    }
  }
  
  return { success: true, results: results };
}

/**
 * シートに移動
 */
function goToSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  if (sheet) {
    spreadsheet.setActiveSheet(sheet);
    return { success: true };
  } else {
    return { success: false, message: 'シートが見つかりません' };
  }
}

/**
 * 統計情報を表示
 */
function showStatistics() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = spreadsheet.getSheetByName('目次');
  
  if (!indexSheet) {
    SpreadsheetApp.getUi().alert('目次シートが見つかりません');
    return;
  }
  
  const lastRow = indexSheet.getLastRow();
  const totalCustomers = Math.floor((lastRow - 1) / 2); // 顧客名簿と顧客データベースの2シートで1顧客
  const totalSheets = spreadsheet.getSheets().length;
  
  const message = `📊 顧客データ統計情報\n\n` +
    `総顧客数: ${totalCustomers}件\n` +
    `総シート数: ${totalSheets}シート\n` +
    `目次エントリ数: ${lastRow - 1}件\n\n` +
    `最終更新: ${new Date()}`;
  
  SpreadsheetApp.getUi().alert(message);
}
