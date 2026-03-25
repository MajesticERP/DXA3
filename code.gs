// ===============================================================
// File: Code.gs (Full Stock Management System with Reporting)
// FINAL VERSION 3.2 (Cleaned & Robust):
// - ลบ const products ออกแล้ว
// - รองรับ Max Stock, รายงานค่าใช้จ่าย, และตัวกรองตามสินค้า
// - (แก้ไข) เพิ่ม Safety Check ใน updateStockManually
// ===============================================================

// --- CONFIGURATION ---
const INVENTORY_SHEET_NAME = 'Inventory';
const MOVEMENTS_SHEET_NAME = 'Movements';

// --- SETUP & WEB APP ---
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⭐ ระบบสต็อก DX-A3')
    .addItem('🚀 เปิด Web App', 'openWebAppDialog')
    .addSeparator()
    .addItem('🛠️ ตรวจสอบ/สร้างชีต (Setup)', 'setupSheets')
    .addToUi();
}

function openWebAppDialog() {
  const url = ScriptApp.getService().getUrl();
  const html = `<p>คลิกที่ลิงก์ด้านล่างเพื่อเปิด Web App:</p><p><a href="${url}" target="_blank">เปิดระบบจัดการสต็อก</a></p>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(300).setHeight(100), 'ลิงก์สำหรับ Web App');
}

/**
 * (ปรับปรุง) ตรวจสอบและสร้างชีต (ไม่ใช้ const products)
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let alertMessage = '';

  // --- Inventory Sheet ---
  const inventorySheet = ss.getSheetByName(INVENTORY_SHEET_NAME);
  const inventoryHeaders = ['NO', 'MATERIAL NO', 'DESCRIPTION', 'UNIT', 'UNIT PRICE', 'TOTAL VALUE', 'CURRENT STOCK', 'MAX STOCK', 'LAST UPDATED'];
  
  if (!inventorySheet) {
    const newInventorySheet = ss.insertSheet(INVENTORY_SHEET_NAME);
    newInventorySheet.appendRow(inventoryHeaders);
    
    // ตั้งค่ารูปแบบตัวเลข
    newInventorySheet.getRange('E:E').setNumberFormat('฿#,##0.00'); // UNIT PRICE
    newInventorySheet.getRange('F:F').setNumberFormat('฿#,##0.00'); // TOTAL VALUE
    newInventorySheet.getRange('I:I').setNumberFormat('yyyy-mm-dd hh:mm:ss'); // LAST UPDATED
    alertMessage += `สร้างชีต ${INVENTORY_SHEET_NAME} ใหม่เรียบร้อย\n`;
  } else {
     // (สำคัญ) กรณีชีตมีอยู่แล้ว ให้อัปเดต Headers ถ้าจำเป็น (เพิ่ม Max Stock)
    const headers = inventorySheet.getRange(1, 1, 1, inventorySheet.getLastColumn()).getValues()[0];
    if (headers.length < 9 || headers[7] !== 'MAX STOCK') {
      inventorySheet.getRange(1, 1, 1, 9).setValues([inventoryHeaders]);
      
      // ตั้งค่าสูตร (ถ้าแถว 2+ ไม่มีสูตร)
      const lastRow = inventorySheet.getLastRow();
      if (lastRow > 1) {
        const formulaRange = inventorySheet.getRange(2, 6, lastRow - 1, 1);
        formulaRange.setFormula('=INDIRECT("R[0]C[-1]",FALSE)*INDIRECT("R[0]C[1]",FALSE)'); // F = E * G
      }
      alertMessage += `อัปเดตโครงสร้างชีต ${INVENTORY_SHEET_NAME} (เพิ่ม MAX STOCK) เรียบร้อย\n`;
    }
  }
  
  // --- Movements Sheet ---
  const movementsSheet = ss.getSheetByName(MOVEMENTS_SHEET_NAME);
  const movementsHeaders = ['TIMESTAMP', 'TYPE', 'DEPARTMENT / NOTE', 'NO', 'MATERIAL NO', 'DESCRIPTION', 'QUANTITY', 'UNIT', 'UNIT PRICE', 'TOTAL VALUE', 'TRANSACTION DATE'];

  if (!movementsSheet) {
    const newMovementsSheet = ss.insertSheet(MOVEMENTS_SHEET_NAME);
    newMovementsSheet.appendRow(movementsHeaders);
    
    // ตั้งค่ารูปแบบตัวเลข
    newMovementsSheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    newMovementsSheet.getRange('I:I').setNumberFormat('฿#,##0.00'); // UNIT PRICE
    newMovementsSheet.getRange('J:J').setNumberFormat('฿#,##0.00'); // TOTAL VALUE
    newMovementsSheet.getRange('K:K').setNumberFormat('yyyy-mm-dd');
    alertMessage += `สร้างชีต ${MOVEMENTS_SHEET_NAME} ใหม่เรียบร้อย\n`;
  }
  
  if (alertMessage === '') {
    alertMessage = 'ชีตทั้งหมดถูกต้องและพร้อมใช้งานแล้ว!';
  }
  
  SpreadsheetApp.getUi().alert(alertMessage);
}


function doGet(e) {
  const action = e.parameter && e.parameter.action;

  // API mode: ?action=xxx
  if (action) {
    let result;
    try {
      switch(action) {
        case 'getInitialData':   result = getInitialData();                        break;
        case 'getInventory':     result = getInventory();                           break;
        case 'getDashboardData': result = getDashboardData();                      break;
        case 'searchEntries':    result = searchEntries(e.parameter.q || '');      break;
        default: result = { status: 'error', message: 'Unknown action: ' + action };
      }
    } catch(err) {
      Logger.log('doGet error: ' + err.toString());
      result = { status: 'error', message: err.message };
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Fallback: serve the HTML (backward compat with GAS web app)
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('DX-A3 Stock Management System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function doPost(e) {
  let params;
  try {
    params = JSON.parse(e.postData.contents);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error', message: 'Invalid request body: ' + err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }

  const action = params.action;
  const data   = params.data;
  let result;
  try {
    switch(action) {
      case 'processTransaction':  result = processTransaction(data);      break;
      case 'updateStockManually': result = updateStockManually(data);     break;
      case 'addNewProduct':       result = addNewProduct(data);           break;
      case 'deleteProduct':       result = deleteProduct(data);           break;
      case 'createReport':        result = createReport(data);            break;
      default: result = { status: 'error', message: 'Unknown action: ' + action };
    }
  } catch(err) {
    Logger.log('doPost error: ' + err.toString());
    result = { status: 'error', message: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- CLIENT-SERVER FUNCTIONS ---

/**
 * ดึงข้อมูลเริ่มต้น (products + inventory) ในการเรียกครั้งเดียว — ลด latency และป้องกัน race condition
 */
function getInitialData() {
  const rows = getInventoryData();
  const products = rows.map(row => ({
    no: row[0],
    materialNo: row[1],
    description: row[2],
    unit: row[3],
    unitPrice: parseFloat(row[4]) || 0,
    maxStock: parseInt(row[7], 10) || 0
  }));
  const inventory = rows.map(row => {
    const price = parseFloat(row[4]) || 0;
    const stock = parseInt(row[6], 10) || 0;
    return {
      no: row[0],
      materialNo: row[1],
      description: row[2],
      unit: row[3],
      unitPrice: price,
      totalValue: price * stock,
      stock: stock,
      maxStock: parseInt(row[7], 10) || 0
    };
  });
  return { products, inventory };
}

/**
 * ดึงข้อมูลสินค้า (รวมราคา และ Max Stock)
 * อ่านจากชีต Inventory (9 คอลัมน์)
 */
function getProducts() {
  return getInventoryData().map(row => ({
    no: row[0],
    materialNo: row[1],
    description: row[2],
    unit: row[3],
    unitPrice: row[4], // [E] UNIT PRICE
    maxStock: row[7]   // [H] MAX STOCK
  }));
}

/**
 * ประมวลผลธุรกรรม (รับเข้า/เบิกจ่าย)
 */
function processTransaction(data) {
  const { type, no, materialNo, description, quantity, unit, department, transactionDate, unitPrice } = data;
  const qty = parseInt(quantity, 10);
  const price = parseFloat(unitPrice) || 0;

  if (isNaN(qty) || qty <= 0) return { status: 'error', message: 'จำนวนต้องเป็นตัวเลขที่มากกว่า 0' };
  
  const quantityChange = (type === 'RECEIVE') ? qty : -qty;
  
  // 1. อัปเดตสต็อก
  invalidateInventoryCache();
  const inventoryResult = updateInventory(materialNo, quantityChange);
  if (inventoryResult.status === 'error') return inventoryResult;

  // 2. (เฉพาะรับเข้า) อัปเดตราคาล่าสุดใน Inventory
  if (type === 'RECEIVE' && price > 0) {
    updateInventoryPrice(materialNo, price);
  }

  // 3. บันทึกการเคลื่อนไหว (11 คอลัมน์)
  const totalValue = qty * price;
  const finalTransactionDate = transactionDate ? new Date(transactionDate) : new Date();
  logMovement(type, department || 'N/A', no, materialNo, description, qty, unit, price, totalValue, finalTransactionDate);

  return { status: 'success' };
}

/**
 * ดึงข้อมูลสต็อกปัจจุบันทั้งหมด
 * อ่านจากชีต Inventory (9 คอลัมน์)
 */
function getInventory() {
  return getInventoryData().map(row => {
    const price = parseFloat(row[4]) || 0;
    const stock = parseInt(row[6], 10) || 0;
    return {
      no: row[0],
      materialNo: row[1],
      description: row[2],
      unit: row[3],
      unitPrice: price,
      totalValue: price * stock,
      stock: stock,
      maxStock: parseInt(row[7], 10) || 0
    };
  });
}

/**
 * ปรับสต็อกด้วยตนเอง
 * (แก้ไข V3.2) เพิ่ม Safety Check
 */
function updateStockManually(data) {
  // *** (V3.2) เพิ่ม Safety Check ป้องกัน data undefined ***
  if (!data) {
    Logger.log('updateStockManually: Error! Received undefined data.');
    return { status: 'error', message: 'ไม่ได้รับข้อมูล (Data is undefined)' };
  }
      
  // *** (อัปเดต V3.1) รับ maxStock ***
  const { materialNo, description, newQuantity, no, maxStock } = data; 
  const newQty = parseInt(newQuantity, 10);
  const newMax = parseInt(maxStock, 10); // *** (ใหม่) ***

  if (isNaN(newQty) || newQty < 0) return { status: 'error', message: 'จำนวนสต็อกใหม่ต้องเป็นตัวเลข 0 หรือมากกว่า' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INVENTORY_SHEET_NAME);
  // อ่าน 9 คอลัมน์
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9);
  const inventoryData = dataRange.getValues();

  for (let i = 0; i < inventoryData.length; i++) {
    if (String(inventoryData[i][1]) === String(materialNo)) { // [B] MATERIAL NO
      const currentRow = i + 2;
      const oldQty = parseInt(inventoryData[i][6], 10) || 0; // [G] CURRENT STOCK
      const unitPrice = parseFloat(inventoryData[i][4]) || 0; // [E] UNIT PRICE

      sheet.getRange(currentRow, 7).setValue(newQty); // อัปเดต [G] CURRENT STOCK
      
      // *** (ใหม่) อัปเดต Max Stock ถ้ามีการส่งค่ามา ***
      if (newMax && newMax > 0) {
        sheet.getRange(currentRow, 8).setValue(newMax); // อัปเดต [H] MAX STOCK
      }

      sheet.getRange(currentRow, 9).setValue(new Date()); // อัปเดต [I] LAST UPDATED
      invalidateInventoryCache();

      // บันทึกการเคลื่อนไหว
      const change = newQty - oldQty;
      const totalValueChange = change * unitPrice;
      logMovement('ADJUST', `ปรับโดยผู้ใช้`, no, materialNo, description, change, `(จาก ${oldQty} เป็น ${newQty})`, unitPrice, totalValueChange, new Date());
      
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'ไม่พบสินค้า' };
}

/**
 * เพิ่มสินค้าใหม่
 */
function addNewProduct(data) {
  // รับ maxStock
  const { no, materialNo, description, unit, unitPrice, maxStock } = data;
  if (!no || !materialNo || !description || !unit) {
    return { status: 'error', message: 'กรุณากรอกข้อมูลให้ครบทุกช่อง (ยกเว้นราคา, สต็อกสูงสุด)' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INVENTORY_SHEET_NAME);
  
  const noColumn = sheet.getRange('A:A').getValues();
  const materialNoColumn = sheet.getRange('B:B').getValues();
  if (noColumn.flat().map(String).includes(String(no)) || materialNoColumn.flat().map(String).includes(String(materialNo))) {
    return { status: 'error', message: 'มี NO หรือ MATERIAL NO นี้อยู่ในระบบแล้ว' };
  }
  
  const price = parseFloat(unitPrice) || 0;
  const max = parseInt(maxStock) || 100; // ค่าเริ่มต้น Max Stock 100
  // อัปเดต appendRow ให้ตรงกับโครงสร้างใหม่ (9 คอลัมน์)
  sheet.appendRow([no, materialNo, description, unit, price, '=INDIRECT("R[0]C[-1]",FALSE)*INDIRECT("R[0]C[1]",FALSE)', 0, max, new Date()]);
  invalidateInventoryCache();
  logMovement('NEW PRODUCT', 'System', no, materialNo, description, 0, unit, price, 0, new Date());
  return { status: 'success', message: 'เพิ่มสินค้าใหม่เรียบร้อยแล้ว' };
}

/**
 * ลบสินค้า (แก้ไขบั๊ก)
 */
function deleteProduct(materialNo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(INVENTORY_SHEET_NAME);
    // อ่าน 9 คอลัมน์
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9);
    const inventoryData = dataRange.getValues();

    for (let i = 0; i < inventoryData.length; i++) {
      if (String(inventoryData[i][1]) === String(materialNo)) { // [B] MATERIAL NO
        const currentStock = parseInt(inventoryData[i][6], 10) || 0; // [G] CURRENT STOCK
        
        if (currentStock > 0) {
          return { status: 'error', message: 'ไม่สามารถลบสินค้าได้ เนื่องจากยังมีสต็อกคงเหลือ (' + currentStock + ' ชิ้น)' };
        }
        
        const productInfo = inventoryData[i];
        sheet.deleteRow(i + 2); // ลบแถว
        invalidateInventoryCache();
        logMovement('DELETE PRODUCT', 'System', productInfo[0], productInfo[1], productInfo[2], 0, productInfo[3], 0, 0, new Date());
        
        return { status: 'success', message: `ลบสินค้า ${productInfo[2]} เรียบร้อยแล้ว` };
      }
    }
    return { status: 'error', message: 'ไม่พบสินค้าที่ต้องการลบ' };
  } catch (e) {
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการลบ: ' + e.message };
  }
}


/**
 * สร้างรายงาน (อัปเกรด: กรองตามสินค้า, สรุปค่าใช้จ่าย)
 */
function createReport(options) {
  try {
    // รับ productFilter
    const { reportType, dataType, startDate, endDate, filterType, productFilter } = options;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let reportName = '';
    let data;

    // 1. จัดการข้อมูลตามประเภทรายงาน
    if (dataType === 'inventory') {
      // --- รายงานสต็อกคงคลัง ---
      const sourceSheet = ss.getSheetByName(INVENTORY_SHEET_NAME);
      if (!sourceSheet) throw new Error(`Sheet "${INVENTORY_SHEET_NAME}" not found.`);
      data = sourceSheet.getDataRange().getValues();
      reportName = `Inventory Report - ${new Date().toLocaleString('th-TH')}`;

    } else if (dataType === 'movements' || dataType === 'expense_summary') {
      // --- รายงานความเคลื่อนไหว หรือ สรุปค่าใช้จ่าย ---
      const sourceSheet = ss.getSheetByName(MOVEMENTS_SHEET_NAME);
      if (!sourceSheet) throw new Error(`Sheet "${MOVEMENTS_SHEET_NAME}" not found.`);
      
      if (sourceSheet.getLastRow() < 2) {
        return { status: 'error', message: 'ไม่พบข้อมูลความเคลื่อนไหว (Movements) ใดๆ' };
      }
      
      data = sourceSheet.getDataRange().getValues();
      const headers = data.shift(); // เก็บ Headers ไว้
      
      // กรองข้อมูล (Filter)
      const start = startDate ? new Date(startDate) : null;
      const end = endDate ? new Date(endDate) : null;
      const productQuery = productFilter ? productFilter.toLowerCase() : null; // ตัวกรองสินค้า

      if(start) start.setHours(0,0,0,0);
      if(end) end.setHours(23,59,59,999);

      data = data.filter(row => {
        const rowDate = new Date(row[10]); // [K] TRANSACTION DATE
        const rowType = row[1]; // [B] TYPE
        const rowDesc = row[5]; // [F] DESCRIPTION
        
        const dateMatch = (!start || rowDate >= start) && (!end || rowDate <= end);
        
        // (ปรับปรุง) ถ้าเป็น expense_summary ให้สนเฉพาะ WITHDRAW
        let typeMatch;
        if (dataType === 'expense_summary') {
          typeMatch = (rowType === 'WITHDRAW');
          // ถ้าเลือก filterType อื่นที่ไม่ใช่ ALL ให้กรองตามนั้นด้วย (เผื่ออนาคต)
          if (filterType !== 'ALL') {
             typeMatch = (rowType === filterType);
          }
        } else {
          typeMatch = (filterType === 'ALL') || (rowType === filterType);
        }

        // เงื่อนไขกรองสินค้า (ใช้ Description)
        const productMatch = (!productQuery) || (rowDesc && rowDesc.toLowerCase().includes(productQuery)); 
        
        return dateMatch && typeMatch && productMatch;
      });

      // 2. ประมวลผลข้อมูลหลังกรอง
      if (dataType === 'expense_summary') {
        // --- ถ้าเป็นรายงานสรุปค่าใช้จ่าย (เฉพาะ WITHDRAW) ---
        reportName = `Expense Summary Report - ${new Date().toLocaleString('th-TH')}`;
        
        // สรุปยอด (Aggregate)
        const summary = {};
        data.forEach(row => {
          const matNo = row[4]; // [E] MATERIAL NO
          const desc = row[5]; // [F] DESCRIPTION
          const unit = row[7]; // [H] UNIT
          const qty = parseInt(row[6], 10) || 0; // [G] QUANTITY
          const val = parseFloat(row[9]) || 0; // [J] TOTAL VALUE
          
          if (!summary[matNo]) {
            summary[matNo] = {
              materialNo: matNo,
              description: desc,
              unit: unit,
              totalQuantity: 0,
              totalValue: 0
            };
          }
          summary[matNo].totalQuantity += qty;
          summary[matNo].totalValue += val;
        });

        // แปลง Object summary เป็น Array สำหรับชีต
        data = [['MATERIAL NO', 'DESCRIPTION', 'UNIT', 'TOTAL QUANTITY (เบิกรวม)', 'TOTAL VALUE (ค่าใช้จ่ายรวม)']];
        for (const key in summary) {
          data.push([
            summary[key].materialNo,
            summary[key].description,
            summary[key].unit,
            summary[key].totalQuantity,
            summary[key].totalValue
          ]);
        }
        
      } else {
        // --- ถ้าเป็นรายงานความเคลื่อนไหว (ปกติ) ---
        reportName = `Movements Report - ${new Date().toLocaleString('th-TH')}`;
        data.unshift(headers); // คืน Headers กลับไป
      }
    }

    if (!data || data.length <= 1) { // <= 1 เพราะอาจจะมีแค่แถว Header
      return { status: 'error', message: 'ไม่พบข้อมูลตามเงื่อนไขที่เลือก' };
    }
    
    // 3. สร้างไฟล์ (เหมือนเดิม)
    const tempSS = SpreadsheetApp.create(reportName);
    const tempSheet = tempSS.getSheets()[0];
    tempSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    SpreadsheetApp.flush();
    const fileId = tempSS.getId();
    
    let blob, mimeType, fileName;

    if (reportType === 'pdf') {
      blob = DriveApp.getFileById(fileId).getAs(MimeType.PDF).setName(`${reportName}.pdf`);
      mimeType = 'application/pdf';
      fileName = `${reportName}.pdf`;
    } else {
      const url = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;
      const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
      blob = response.getBlob().setName(`${reportName}.xlsx`);
      mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      fileName = `${reportName}.xlsx`;
    }
    
    DriveApp.getFileById(fileId).setTrashed(true);

    return {
      status: 'success',
      blob: Utilities.base64Encode(blob.getBytes()),
      fileName: fileName,
      mimeType: mimeType
    };

  } catch (e) {
    Logger.log(e.toString());
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการสร้างรายงาน: ' + e.message };
  }
}

/**
 * ค้นหาประวัติ
 * อ่านจาก Movements (11 คอลัมน์)
 */
function searchEntries(searchQuery) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MOVEMENTS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
  
  if (!searchQuery) {
    // ส่ง 20 รายการล่าสุด (ถ้ามี)
    return data.slice(-20).map(mapMovementRow).reverse();
  }

  const query = searchQuery.toLowerCase();
  const results = data.map(mapMovementRow).filter(item => 
    item.no.toString().toLowerCase().includes(query) ||
    item.materialNo.toLowerCase().includes(query) ||
    item.description.toLowerCase().includes(query) ||
    item.department.toLowerCase().includes(query)
  );
  
  return results.reverse();
}


// --- HELPER & DASHBOARD FUNCTIONS ---

/**
 * อัปเดตสต็อก (Helper)
 * อ่าน/เขียน Inventory (9 คอลัมน์)
 */
function updateInventory(materialNo, quantityChange) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INVENTORY_SHEET_NAME);
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9);
  const data = dataRange.getValues();
  
  let itemFound = false;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1]) === String(materialNo)) { // [B] MATERIAL NO
      const currentRow = i + 2;
      const currentStock = parseInt(data[i][6], 10) || 0; // [G] CURRENT STOCK
      const newStock = currentStock + quantityChange;

      if (newStock < 0) {
        return { status: 'error', message: `สต็อกไม่พอ! เหลืออยู่ ${currentStock} ชิ้น` };
      }
      
      sheet.getRange(currentRow, 7).setValue(newStock); // อัปเดต [G] CURRENT STOCK
      sheet.getRange(currentRow, 9).setValue(new Date()); // อัปเดต [I] LAST UPDATED
      // [F] TOTAL VALUE จะอัปเดตอัตโนมัติ
      itemFound = true;
      break;
    }
  }

  if (!itemFound) return { status: 'error', message: `ไม่พบสินค้า ${materialNo} ในคลัง` };
  return { status: 'success' };
}

/**
 * อัปเดตราคาใน Inventory (Helper)
 */
function updateInventoryPrice(materialNo, newPrice) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INVENTORY_SHEET_NAME);
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5); // ต้องการแค่ 5 คอลัมน์แรก
  const data = dataRange.getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1]) === String(materialNo)) { // [B] MATERIAL NO
      const currentRow = i + 2;
      sheet.getRange(currentRow, 5).setValue(newPrice); // อัปเดต [E] UNIT PRICE
      invalidateInventoryCache();
      break;
    }
  }
}


/**
 * บันทึกการเคลื่อนไหว (Helper)
 * เขียนไป Movements (11 คอลัมน์)
 */
function logMovement(type, department, no, materialNo, description, quantity, unit, unitPrice, totalValue, transactionDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MOVEMENTS_SHEET_NAME);
  sheet.appendRow([new Date(), type, department, no, materialNo, description, quantity, unit, unitPrice, totalValue, transactionDate]);
}

/**
 * ดึงข้อมูล Movements ทั้งหมด (Helper)
 */
function getMovementsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MOVEMENTS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  // อ่าน 11 คอลัมน์
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
}

/**
 * ล้าง Cache Inventory (เรียกหลังทุก write operation)
 */
function invalidateInventoryCache() {
  try { CacheService.getScriptCache().remove(CACHE_KEY_INVENTORY); } catch(e) {}
}

/**
 * ดึงข้อมูล Inventory ทั้งหมด (Helper) — ใช้ CacheService เพื่อลด latency
 */
function getInventoryData() {
  try {
    const cached = CacheService.getScriptCache().get(CACHE_KEY_INVENTORY);
    if (cached) return JSON.parse(cached);
  } catch(e) {
    Logger.log('Cache read error: ' + e);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INVENTORY_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();

  try {
    CacheService.getScriptCache().put(CACHE_KEY_INVENTORY, JSON.stringify(data), CACHE_DURATION);
  } catch(e) {
    Logger.log('Cache write error: ' + e);
  }

  return data;
}

/**
 * ดึงข้อมูลสำหรับแดชบอร์ด
 */
function getDashboardData() {
  const movements = getMovementsData();
  const inventory = getInventoryData();
  
  // คำนวณมูลค่าสต็อกรวม
  const totalInventoryValue = inventory.reduce((acc, row) => {
    return acc + (parseFloat(row[5]) || 0); // [F] TOTAL VALUE
  }, 0);

  // (เพิ่มใหม่) คำนวณสินค้าที่ต้องเติม (น้อยกว่า 20% ของ Max Stock)
  const restockItems = inventory.map(item => {
    const stock = parseInt(item[6], 10) || 0;     // [G] CURRENT STOCK
    const maxStock = parseInt(item[7], 10) || 0; // [H] MAX STOCK
    const percent = (maxStock > 0) ? (stock / maxStock) * 100 : 0;
    return {
      description: item[2],
      stock: stock,
      maxStock: maxStock,
      percent: percent
    }
  }).filter(item => item.maxStock > 0 && item.percent < 20) // กรองเฉพาะที่น้อยกว่า 20%
    .sort((a, b) => a.percent - b.percent) // เรียงจากน้อยสุดไปมากสุด
    .slice(0, 10); // เอา 10 รายการ

  return {
    dailyWithdrawals: calculateDailyWithdrawals(movements),
    topWithdrawnItems: calculateTopItems(movements, 'WITHDRAW'),
    slowMovingStock: getSlowMovingItems(inventory),
    inventoryLevels: getInventoryStockLevels(inventory),
    totalInventoryValue: totalInventoryValue,
    restockItems: restockItems // ส่งข้อมูลใหม่ไปหน้าเว็บ
  };
}

function calculateDailyWithdrawals(movements) {
  const dailyData = {};
  const today = new Date();
  
  movements.forEach(row => {
    const type = row[1];
    if (type === 'WITHDRAW') {
      const timestamp = new Date(row[0]);
      const diffDays = Math.floor((today - timestamp) / (1000 * 60 * 60 * 24));

      if (diffDays < 7) {
        const day = timestamp.toLocaleDateString('th-TH', { weekday: 'short' });
        const quantity = parseInt(row[6], 10); // [G] QUANTITY
        dailyData[day] = (dailyData[day] || 0) + quantity;
      }
    }
  });
  
  const labels = [];
  for(let i = 6; i >= 0; i--) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    labels.push(d.toLocaleDateString('th-TH', { weekday: 'short' }));
  }
  
  const values = labels.map(label => dailyData[label] || 0);

  return { labels, values };
}

function calculateTopItems(movements, type) {
  const itemCounts = {};
  movements.forEach(row => {
    if (row[1] === type) {
      const description = row[5];
      const quantity = parseInt(row[6], 10); // [G] QUANTITY
      itemCounts[description] = (itemCounts[description] || 0) + quantity;
    }
  });

  const sortedItems = Object.entries(itemCounts).sort(([, a], [, b]) => b - a);
  const top5 = sortedItems.slice(0, 5);

  return {
    labels: top5.map(item => item[0]),
    values: top5.map(item => item[1])
  };
}

function getSlowMovingItems(inventory) {
  const today = new Date();
  const slowItems = inventory.map(item => {
    const lastUpdated = new Date(item[8]); // [I] LAST UPDATED
    const diffDays = Math.floor((today - lastUpdated) / (1000 * 60 * 60 * 24));
    return {
      description: item[2], // [C] DESCRIPTION
      stock: item[6],       // [G] CURRENT STOCK
      daysSinceUpdate: diffDays
    };
  }).filter(item => item.daysSinceUpdate > 30 && item.stock > 0) // เพิ่มเงื่อนไข stock > 0
    .sort((a, b) => b.daysSinceUpdate - a.daysSinceUpdate)
    .slice(0, 10);

  return slowItems;
}

function getInventoryStockLevels(inventory) {
  const itemsWithStock = inventory
    .filter(item => parseInt(item[6], 10) > 0) // [G] CURRENT STOCK
    .sort((a, b) => parseInt(b[6], 10) - parseInt(a[6], 10))
    .slice(0, 15);

  return {
    labels: itemsWithStock.map(item => item[2]),
    values: itemsWithStock.map(item => item[6])
  };
}

/**
 * แปลงแถวข้อมูล Movements เป็น Object
 * (11 คอลัมน์)
 */
function mapMovementRow(row) {
  return {
    timestamp: row[0] ? new Date(row[0]).toLocaleString('th-TH') : 'N/A',
    type: row[1] || '',
    department: row[2] || '',
    no: row[3] || '',
    materialNo: row[4] || '',
    description: row[5] || '',
    quantity: row[6] || '',
    unit: row[7] || '',
    unitPrice: row[8] || 0,
    totalValue: row[9] || 0,
    transactionDate: row[10] ? new Date(row[10]).toLocaleDateString('th-TH') : 'N/A'
  };
}

