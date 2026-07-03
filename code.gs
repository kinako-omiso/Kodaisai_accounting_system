// URLが https://docs.google.com/spreadsheets/d/ABCDEFG123456/edit の場合、ABCDEFG123456 の部分
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');

const SHEET_INVENTORY = '在庫管理';
const SHEET_ORDERS = '注文履歴';

// Webアプリを開いたときに呼ばれる関数2
function doGet(e) {
  if (e.parameter && e.parameter.page === 'staff') {
    return HtmlService.createHtmlOutputFromFile('staff')
      .setTitle('スタッフ用管理画面')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('工大祭 注文フォーム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

// スプレッドシートから在庫一覧を取得する
function getProducts() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const invSheet = ss.getSheetByName(SHEET_INVENTORY);
  const invData = invSheet.getDataRange().getValues();
  
  const productsMap = {};
  for (let i = 1; i < invData.length; i++) {
    const row = invData[i];
    if (row[0]) {
      productsMap[row[2]] = {
        id: row[0],
        category: row[1],
        name: row[2],
        price: row[3],
        stock: row[4],
        image: row[6] || '',
        pendingQty: 0
      };
    }
  }

  const orderSheet = ss.getSheetByName(SHEET_ORDERS);
  const orderData = orderSheet.getDataRange().getValues();
  
  for (let i = 1; i < orderData.length; i++) {
    const row = orderData[i];
    const detailsStr = row[3];
    const foodStatus = row[7];
    const goodsStatus = row[9];
    
    if (!detailsStr) continue;
    
    const items = detailsStr.split(', ');
    items.forEach(itemStr => {
      const lastXIndex = itemStr.lastIndexOf('x');
      if (lastXIndex !== -1) {
        const name = itemStr.substring(0, lastXIndex);
        const qty = parseInt(itemStr.substring(lastXIndex + 1));
        
        if (productsMap[name]) {
          const cat = productsMap[name].category;
          if (cat === '飲食' && foodStatus === '未') {
            productsMap[name].pendingQty += qty;
          } else if (cat === 'グッズ' && goodsStatus === '未') {
            productsMap[name].pendingQty += qty;
          }
        }
      }
    });
  }

  const products = [];
  for (const name in productsMap) {
    const p = productsMap[name];
    const availableStock = p.stock - p.pendingQty;
    products.push({
      id: p.id,
      category: p.category,
      name: p.name,
      price: p.price,
      stock: availableStock > 0 ? availableStock : 0,
      image: p.image
    });
  }
  
  return products;
}



function placeOrder(orderData) {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const stockSheet = ss.getSheetByName(SHEET_INVENTORY);
  const stockData = stockSheet.getDataRange().getValues();
  
  const validNames = stockData.slice(1).map(row => String(row[2]));

  for (let i = 0; i < orderData.items.length; i++) {
    const item = orderData.items[i];

    if (item.qty <= 0) {
      throw new Error("0以下の注文数が送信されました。");
    }

    if (!validNames.includes(String(item.name))) {
      throw new Error("【存在しない商品名が送信されました。");
    }
  }
  
  const currentProducts = getProducts();
  const currentStockMap = {};
  currentProducts.forEach(p => {
    currentStockMap[p.name] = p.stock;
  });

  let isSoldOutError = false;
  orderData.items.forEach(item => {
    const latestStock = currentStockMap[item.name] || 0;
    if (item.qty > latestStock) {
      isSoldOutError = true; // 注文数が最新の残数を超えていたらエラー
    }
  });

  if (isSoldOutError) {
    return { error: 'sold_out' }; 
  }
 

  const sheet = ss.getSheetByName(SHEET_ORDERS);
  const lastRow = sheet.getLastRow();
  let nextOrderId = 1;
  if (lastRow > 1) {
    const lastId = sheet.getRange(lastRow, 1).getValue();
    nextOrderId = Number(lastId) + 1;
  }
  
  let hasFood = false;
  let hasGoods = false;
  let orderTextArr = [];
  
  orderData.items.forEach(item => {
    if(item.category === '飲食') hasFood = true;
    if(item.category === 'グッズ') hasGoods = true;
    orderTextArr.push(`${item.name}x${item.qty}`);
  });
  
  const orderText = orderTextArr.join(', ');
  const actualFoodId = hasFood ? 1000 + nextOrderId : '-';
  const actualGoodsId = hasGoods ? 9000 + nextOrderId : '-';
  
  const now = new Date();
  const timeString = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm');
  
  sheet.appendRow([
    nextOrderId,
    timeString,
    orderData.nickname,
    orderText,
    orderData.total,
    '未',
    actualFoodId,
    hasFood ? '未' : '-',
    actualGoodsId,
    hasGoods ? '未' : '-'
  ]);
  
  return {
    orderId: nextOrderId,
    foodId: actualFoodId,
    goodsId: actualGoodsId,
    total: orderData.total
  };
}

//スキャンした注文IDから情報を取得
function getOrderInfo(orderId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_ORDERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == orderId) { // 注文IDが一致
      return {
        row: i + 1,
        id: data[i][0],
        nickname: data[i][2],
        details: data[i][3],
        total: data[i][4],
        paymentStatus: data[i][5]
      };
    }
  }
  return null; // 見つからない場合
}

function completePayment(orderId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_ORDERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == orderId) {
      sheet.getRange(i + 1, 6).setValue('済'); // F列（会計状態）を済に
      return true;
    }
  }
  return false;
}

//未提供のリストを取得
function getPendingOrders(type) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_ORDERS);
  const data = sheet.getDataRange().getValues();
  
  const invSheet = ss.getSheetByName(SHEET_INVENTORY);
  const invData = invSheet.getDataRange().getValues();
  const categoryMap = {};
  for(let i=1; i<invData.length; i++) {
     categoryMap[invData[i][2]] = invData[i][1]; // { "小籠包": "飲食", "アクスタ1": "グッズ" }
  }
  
  const pendingList = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const paymentStatus = row[5];
    const foodStatus = row[7];
    const goodsStatus = row[9];
    
    // 会計済みのものだけ表示
    if (paymentStatus === '済') {
      const detailsStr = row[3];
      const items = detailsStr.split(', ');
      const targetItems = [];
      
      // 注文内容から、現在のタブ(type)に一致する商品だけを抽出
      items.forEach(itemStr => {
        const lastXIndex = itemStr.lastIndexOf('x');
        if (lastXIndex !== -1) {
          const name = itemStr.substring(0, lastXIndex);
          const cat = categoryMap[name];
          if (type === 'food' && cat === '飲食') targetItems.push(itemStr);
          if (type === 'goods' && cat === 'グッズ') targetItems.push(itemStr);
        }
      });
      
      const filteredDetails = targetItems.join(', ');
      
      // 対象の商品が含まれている場合のみリストに追加
      if (type === 'food' && foodStatus === '未' && targetItems.length > 0) {
        pendingList.push({ row: i + 1, displayId: row[6], nickname: row[2], details: filteredDetails });
      } else if (type === 'goods' && goodsStatus === '未' && targetItems.length > 0) {
        pendingList.push({ row: i + 1, displayId: row[8], nickname: row[2], details: filteredDetails });
      }
    }
  }
  return pendingList;
}

function completeDelivery(row, type) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const orderSheet = ss.getSheetByName(SHEET_ORDERS);
  const invSheet = ss.getSheetByName(SHEET_INVENTORY);
  
  if (type === 'food') orderSheet.getRange(row, 8).setValue('済');
  if (type === 'goods') orderSheet.getRange(row, 10).setValue('済');

  const orderDetails = orderSheet.getRange(row, 4).getValue();
  const items = orderDetails.split(', ');
  const invData = invSheet.getDataRange().getValues();
  
  items.forEach(itemStr => {
    const lastXIndex = itemStr.lastIndexOf('x');
    if (lastXIndex !== -1) {
      const name = itemStr.substring(0, lastXIndex);
      const qty = parseInt(itemStr.substring(lastXIndex + 1));
      
      for (let i = 1; i < invData.length; i++) {
        const invCategory = invData[i][1];
        const invName = invData[i][2];
      
        let isTarget = false;
        if (type === 'food' && invCategory === '飲食') isTarget = true;
        if (type === 'goods' && invCategory === 'グッズ') isTarget = true;
        
        if (isTarget && invName === name) {
          const currentStock = invSheet.getRange(i + 1, 5).getValue();
          invSheet.getRange(i + 1, 5).setValue(currentStock - qty);
          break;
        }
      }
    }
  });
  return true;
}

function certification(password) {
  const truePassword = PropertiesService.getScriptProperties().getProperty('STAFF_PASSWORD');
  if (password === truePassword) {
    return HtmlService.createHtmlOutputFromFile('staff_ui').getContent();;  
  } else {
    return null; 
  }
}