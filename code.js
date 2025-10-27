const ss = SpreadsheetApp.getActiveSpreadsheet();
const salesSheet = ss.getSheetByName('売上');
const expenseSheet = ss.getSheetByName('支出');
const menuSheet = ss.getSheetByName('メニュー');
const settingsSheet = ss.getSheetByName('設定');

function doGet(e) {
  return ContentService.createTextOutput("文化祭イカ焼き会計アプリのバックエンドへようこそ！");
}

function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    let result = {};

    switch (action) {
      case 'recordSale':
        result = recordSale(request.payload);
        break;
      // ★追加: パスワード認証と価格更新のアクション
      case 'verifyPassword':
        result = verifyPassword(request.payload);
        break;
      case 'updatePrice':
        result = updatePrice(request.payload);
        break;
      case 'getDataForDate':
        result = getDataForDate(request.payload);
        break;
      default:
        throw new Error('無効なアクションです。');
    }
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function recordSale(payload) { /* 変更なし */ }

// ★追加: パスワードを認証する関数
function verifyPassword(payload) {
    const settingsData = settingsSheet.getDataRange().getValues();
    let correctPassword = '';
    // 設定シートから'admin_password'を探す
    for (let i = 0; i < settingsData.length; i++) {
        if (settingsData[i][0] === 'admin_password') {
            correctPassword = settingsData[i][1];
            break;
        }
    }
    // パスワードが一致するかどうかをtrue/falseで返す
    return { authenticated: (correctPassword !== '' && payload.password === correctPassword) };
}

// ★追加: 価格を更新する関数
function updatePrice(payload) {
  const menuData = menuSheet.getDataRange().getValues();
  // メニューシートを1行目から探し、商品名が一致したら価格(3列目)を更新
  for (let i = 1; i < menuData.length; i++) {
    if (menuData[i][1] === payload.name) {
      menuSheet.getRange(i + 1, 3).setValue(payload.newPrice);
      return { message: `${payload.name}の価格を${payload.newPrice}円に更新しました。` };
    }
  }
  throw new Error('該当する商品が見つかりません。');
}

function getDataForDate(payload) { /* 変更なし */ }


// --- 変更のない関数の完全なコード ---
function recordSale(payload) {
  const now = new Date();
  const date = Utilities.formatDate(now, "JST", "yyyy/MM/dd");
  const timestamp = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm:ss");
  const lastRow = salesSheet.getLastRow();
  let newAccountingId;
  if (lastRow < 2) { newAccountingId = 101; }
  else { newAccountingId = Number(salesSheet.getRange(lastRow, 1).getValue()) + 1; }
  payload.items.forEach(item => {
    const subtotal = item.quantity * item.price;
    salesSheet.appendRow([newAccountingId, timestamp, date, item.name, item.quantity, item.price, subtotal]);
  });
  return { message: '売上を記録しました。', accountingId: newAccountingId };
}
function getDataForDate(payload) {
    const targetDate = payload.date;
    const menuData = menuSheet.getDataRange().getValues().slice(1)
        .map(row => ({ name: row[1], price: row[2], status: row[3] }));
    let salesHistory = [];
    if (salesSheet.getLastRow() > 1) {
        salesHistory = salesSheet.getDataRange().getValues().slice(1)
            .filter(row => row[2].toString() === targetDate)
            .map(row => ({ id: row[0], timestamp: row[1], name: row[3], quantity: row[4], price: row[5], subtotal: row[6] }));
    }
    return { menu: menuData, sales: salesHistory };
}