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
      case 'recordExpense':
        result = recordExpense(request.payload);
        break;
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

function recordExpense(payload) {
    const lastRow = expenseSheet.getLastRow();
    const newExpenseId = (lastRow < 2) ? 1 : Number(expenseSheet.getRange(lastRow, 1).getValue()) + 1;
    expenseSheet.appendRow([newExpenseId, payload.date, payload.name, payload.amount, '']);
    return { message: '支出を記録しました。' };
}

function verifyPassword(payload) {
    const settingsData = settingsSheet.getDataRange().getValues();
    let correctPassword = '';
    for (let i = 0; i < settingsData.length; i++) {
        if (settingsData[i][0] === 'admin_password') {
            correctPassword = settingsData[i][1];
            break;
        }
    }
    return { authenticated: (correctPassword !== '' && payload.password.toString() === correctPassword.toString()) };
}

function updatePrice(payload) {
  const menuData = menuSheet.getDataRange().getValues();
  for (let i = 1; i < menuData.length; i++) {
    if (menuData[i][1] === payload.name) {
      menuSheet.getRange(i + 1, 3).setValue(payload.newPrice);
      return { message: `${payload.name}の価格を${payload.newPrice}円に更新しました。` };
    }
  }
  throw new Error('該当する商品が見つかりません。');
}

function getDataForDate(payload) {
    const targetDate = payload.date;

    const menuData = menuSheet.getDataRange().getValues().slice(1)
        .map(row => ({ name: row[1], price: row[2], status: row[3] }));

    let salesHistory = [];
    if (salesSheet.getLastRow() > 1) {
        salesHistory = salesSheet.getDataRange().getValues().slice(1)
            .filter(row => {
                // 日付セルが空でないことを確認
                if (row[2]) {
                    // スプレッドシートの日付を 'yyyy/MM/dd' 形式の文字列に変換して比較
                    const sheetDate = Utilities.formatDate(new Date(row[2]), "JST", "yyyy/MM/dd");
                    return sheetDate === targetDate;
                }
                return false;
            })
            .map(row => ({ id: row[0], timestamp: row[1], name: row[3], quantity: row[4], price: row[5], subtotal: row[6] }));
    }

    // ★★★ 修正ポイント ★★★
    let expenseHistory = [];
    if (expenseSheet.getLastRow() > 1) {
        expenseHistory = expenseSheet.getDataRange().getValues().slice(1)
            .filter(row => {
                // 日付セルが空でないことを確認
                if (row[1]) {
                    // スプレッドシートの日付を 'yyyy/MM/dd' 形式の文字列に変換して比較
                    const sheetDate = Utilities.formatDate(new Date(row[1]), "JST", "yyyy/MM/dd");
                    return sheetDate === targetDate;
                }
                return false;
            })
            .map(row => ({ id: row[0], date: row[1], name: row[2], amount: row[3] }));
    }

    return { menu: menuData, sales: salesHistory, expenses: expenseHistory };
}