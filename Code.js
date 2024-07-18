function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function getProductsItems() {
  const sheet = SpreadsheetApp.openById('1nSI3nKWilvTeymJN2H_rT2_wfOwLnyVr2EOgvSKak04').getSheetByName('รายการ');
  const data = sheet.getDataRange().getValues();
  data.shift(); // Remove headers
  const lowStockThreshold = 1; // Define your own threshold
  const filteredData = data.filter(row => row[7] >= lowStockThreshold);
  return JSON.parse(JSON.stringify(filteredData));
}

function getLowStockItems() {
  const sheet = SpreadsheetApp.openById('1nSI3nKWilvTeymJN2H_rT2_wfOwLnyVr2EOgvSKak04').getSheetByName('รายการ');
  const data = sheet.getDataRange().getValues();
  data.shift(); // Remove headers
  const lowStockThreshold = 5; // Define your own threshold
  const filteredData = data.filter(row => row[7] < lowStockThreshold && row[7] >= 1);
  return JSON.parse(JSON.stringify(filteredData));
}

function getOutStockItems() {
  const sheet = SpreadsheetApp.openById('1nSI3nKWilvTeymJN2H_rT2_wfOwLnyVr2EOgvSKak04').getSheetByName('รายการ');
  const data = sheet.getDataRange().getValues();
  data.shift(); // Remove headers
  const outStockThreshold = 1; // Define your own threshold
  const filteredData = data.filter(row => row[7] < outStockThreshold || isNaN(row[7]));
  return JSON.parse(JSON.stringify(filteredData));
}

function getCategoryOptions() {
  const sheet = SpreadsheetApp.openById('1nSI3nKWilvTeymJN2H_rT2_wfOwLnyVr2EOgvSKak04').getSheetByName('หมวดหมู่');
  const data = sheet.getRange('B2:B' + sheet.getLastRow()).getValues(); // Column B is the "ชื่อหมวดหมู่่" column
  const categories = [...new Set(data.flat().filter(category => category))]; // Remove duplicates and empty values
  console.log(categories)
  return categories;
}

function addProductToSheet(productID, productName, category, size, unit, costPrice, salePrice, stockAmount, recordDate, manufacturer, note) {
  const sheet = SpreadsheetApp.openById('1nSI3nKWilvTeymJN2H_rT2_wfOwLnyVr2EOgvSKak04').getSheetByName('รายการ');
  sheet.appendRow([productID, productName, category, size, unit, costPrice, salePrice, stockAmount, recordDate, manufacturer, note]);
}

function getNewOrderID() {
  const sheet = SpreadsheetApp.openById('1nSI3nKWilvTeymJN2H_rT2_wfOwLnyVr2EOgvSKak04').getSheetByName('สั่งซื้อ');
  const data = sheet.getDataRange().getValues();
  const lastOrderID = data.length > 1 ? parseInt(data[data.length - 1][1]) : 0;
  return lastOrderID + 1;
}

function getNewProductID() {
  const sheet = SpreadsheetApp.openById('1nSI3nKWilvTeymJN2H_rT2_wfOwLnyVr2EOgvSKak04').getSheetByName('รายการ');
  const data = sheet.getDataRange().getValues();
  const lastProductID = data.length > 1 ? parseInt(data[data.length - 1][0]) : 0;
  return lastProductID + 1;
}

function getCustomers() {
  const sheet = SpreadsheetApp.openById('1nSI3nKWilvTeymJN2H_rT2_wfOwLnyVr2EOgvSKak04').getSheetByName('ลูกค้า');
  const data = sheet.getDataRange().getValues();
  data.shift(); // Remove headers
  return data;
}

function getProductByName(productName) {
  const sheet = SpreadsheetApp.openById('1nSI3nKWilvTeymJN2H_rT2_wfOwLnyVr2EOgvSKak04').getSheetByName('รายการ');
  const data = sheet.getDataRange().getValues();
  const product = data.find(row => row[2] === productName);
  return product || [];
}

function convertNumberToThai(number) {
  const thaiNumbers = ["ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"];
  const units = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน", "ล้าน"];
  let bahtText = "";
  const digits = number.toString().split("").reverse();
  
  digits.forEach((digit, index) => {
    const num = parseInt(digit);
    if (num !== 0) {
      bahtText = thaiNumbers[num] + units[index] + bahtText;
    }
  });
  
  return bahtText + "บาทถ้วน";
}

function confirmSale(orderID, saleType, customerName, customerTel, totalPrice, saleItems) {
  const sheet = SpreadsheetApp.openById('1nSI3nKWilvTeymJN2H_rT2_wfOwLnyVr2EOgvSKak04').getSheetByName('สั่งซื้อ');
  const timestamp = new Date();
  saleItems.forEach(item => {
    sheet.appendRow([timestamp, orderID, saleType, customerName, item.productID, item.amount, item.price, totalPrice]);
  });
}

function getSummaryData(period) {
  const sheet = SpreadsheetApp.openById('1nSI3nKWilvTeymJN2H_rT2_wfOwLnyVr2EOgvSKak04').getSheetByName('สั่งซื้อ');
  const data = sheet.getDataRange().getValues();
  data.shift(); // Remove headers
  
  const summary = {};
  data.forEach(row => {
    const date = new Date(row[0]);
    let key;
    if (period === 'daily') {
      key = date.toLocaleDateString('th-TH');
    } else if (period === 'monthly') {
      key = `${date.getFullYear()}-${date.getMonth() + 1}`;
    } else if (period === 'yearly') {
      key = date.getFullYear();
    }
    if (!summary[key]) {
      summary[key] = 0;
    }
    summary[key] += parseFloat(row[7]);
  });
  
  const labels = Object.keys(summary);
  const values = Object.values(summary);
  return { labels, values };
}
