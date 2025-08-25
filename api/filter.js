import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { Readable } from 'stream';

// Improved multipart parser with better binary handling
function parseMultipartData(buffer, boundary) {
  const files = {};
  const boundaryBuffer = Buffer.from(`--${boundary}`, 'utf8');
  const parts = [];
  
  let start = 0;
  let boundaryIndex = buffer.indexOf(boundaryBuffer, start);
  
  while (boundaryIndex !== -1) {
    if (start !== boundaryIndex) {
      parts.push(buffer.slice(start, boundaryIndex));
    }
    start = boundaryIndex + boundaryBuffer.length;
    boundaryIndex = buffer.indexOf(boundaryBuffer, start);
  }
  
  for (const part of parts) {
    if (part.length < 10) continue;
    
    const partStr = part.toString('binary', 0, Math.min(part.length, 500));
    if (!partStr.includes('Content-Disposition')) continue;
    
    const headerEndIndex = part.indexOf('\r\n\r\n');
    if (headerEndIndex === -1) continue;
    
    const headers = part.slice(0, headerEndIndex).toString('utf8');
    const nameMatch = headers.match(/name="([^"]+)"/);
    const filenameMatch = headers.match(/filename="([^"]+)"/);
    
    if (!nameMatch || !filenameMatch) continue;
    
    const name = nameMatch[1];
    const filename = filenameMatch[1];
    const fileData = part.slice(headerEndIndex + 4, part.length - 2); // Remove final \r\n
    files[name] = {
      buffer: fileData,
      filename: filename
    };
  }
  
  return files;
}

// Column mappings and other constants (copied from server/utils/filterExcel.js)
const COLUMN_MAPPINGS = {
  'Order Number': 'Order #',
  'Flat Number': 'Flat #',
  'Customer Mobile Number': 'Mobile No',
  'Confirmed Order': 'Cnf',
  'Product Name': 'Product Name',
  'Item Count': 'Qty',
  'Rate': 'Price',
  'Item total': 'I Tot',
  'Total Items': 'Total Items',
  'Payment Mode': 'Payment Mode',
  'Payment Status': 'Payment Status',
  'Total Amount': 'T Amt'
};

const COLUMN_WIDTHS = {
  'Order #': 7,
  'Flat #': 7,
  'Mobile No': 16,
  'Cnf': 2,
  'Product Name': 35,
  'Qty': 2.5,
  'Price': 5.5,
  'I Tot': 5,
  'Total Items': 6,
  'Payment Mode': 5.75,
  'Payment Status': 5.75,
  'T Amt': 6,
};

const COLUMN_WIDTHS_SHEET2 = {
  ...COLUMN_WIDTHS,
  'Catalogue Group': 20,
  'Tax %': 8,
  'Tax Amount': 10
};

// Utility functions (copied from server/utils/filterExcel.js)
const normalizeCache = new Map();

function normalizeFlatString(str) {
  const key = String(str || '');
  if (normalizeCache.has(key)) {
    return normalizeCache.get(key);
  }
  
  const normalized = key
    .replace(/[^A-Z0-9]/gi, '')
    .toUpperCase();
  
  normalizeCache.set(key, normalized);
  return normalized;
}

function extractNumberFromAddress(address) {
  const addressStr = String(address || '').trim();
  const normalizedStr = normalizeFlatString(addressStr);
  const match = normalizedStr.match(/\d+/);
  return match ? match[0] : '';
}

const flatParseCache = new Map();

function parseFlatNumber(flatNo) {
  const str = String(flatNo || '').trim();
  
  if (flatParseCache.has(str)) {
    return flatParseCache.get(str);
  }
  
  if (!str) {
    const result = { tower: '', floor: 0, apt: 0, isEmpty: true };
    flatParseCache.set(str, result);
    return result;
  }
  
  const match = str.match(/^([A-Z]+)(\d+)$/i);
  if (!match) {
    const result = { tower: str, floor: 0, apt: 0, isEmpty: !str };
    flatParseCache.set(str, result);
    return result;
  }
  
  const tower = match[1].toUpperCase();
  const numberPart = match[2];
  
  let floor, apt;
  if (numberPart.length === 3) {
    floor = parseInt(numberPart.substring(0, 1));
    apt = parseInt(numberPart.substring(1));
  } else if (numberPart.length === 4) {
    floor = parseInt(numberPart.substring(0, 2));
    apt = parseInt(numberPart.substring(2));
  } else {
    floor = parseInt(numberPart);
    apt = 0;
  }
  
  const result = { tower, floor, apt, isEmpty: false };
  flatParseCache.set(str, result);
  return result;
}

function extractSheetData(sheet) {
  const data = [];
  const headers = [];
  
  const headerRow = sheet.getRow(1);
  headerRow.eachCell((cell, colNumber) => {
    headers[colNumber] = cell.value;
  });
  
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    
    const rowData = {};
    row.eachCell((cell, colNumber) => {
      const header = headers[colNumber];
      if (header) {
        rowData[header] = cell.value || '';
      }
    });
    data.push(rowData);
  });
  
  return data;
}

function processOrders(data) {
  const orderGroups = new Map();
  const orderTotals = new Map();
  const orderItemTotals = new Map();
  
  data.forEach(row => {
    const orderNum = row['Order Number'];
    
    if (!orderGroups.has(orderNum)) {
      orderGroups.set(orderNum, []);
      orderTotals.set(orderNum, 0);
      orderItemTotals.set(orderNum, 0);
    }
    
    orderGroups.get(orderNum).push(row);
    
    const count = parseFloat(row['Item Count']) || 0;
    const discountedPrice = parseFloat(row['Discounted Price']) || 0;
    const regularPrice = parseFloat(row['Price']) || 0;
    const price = discountedPrice || regularPrice;
    
    orderTotals.set(orderNum, orderTotals.get(orderNum) + (count * price));
    orderItemTotals.set(orderNum, orderItemTotals.get(orderNum) + count);
  });
  
  for (const [orderNum, total] of orderTotals) {
    orderTotals.set(orderNum, Math.round(total * 100) / 100);
  }
  
  return { orderGroups, orderTotals, orderItemTotals };
}

function groupAndSort(orders, moveEmptyToBottom = false) {
  const groups = new Map();
  
  orders.forEach(row => {
    const orderNum = row['Order Number'];
    if (!groups.has(orderNum)) {
      groups.set(orderNum, []);
    }
    groups.get(orderNum).push(row);
  });
  
  const sortedGroups = Array.from(groups.entries()).sort((a, b) => {
    const flatA = parseFlatNumber(a[1][0]['Flat Number']);
    const flatB = parseFlatNumber(b[1][0]['Flat Number']);
    
    if (moveEmptyToBottom) {
      if (flatA.isEmpty && !flatB.isEmpty) return 1;
      if (!flatA.isEmpty && flatB.isEmpty) return -1;
      if (flatA.isEmpty && flatB.isEmpty) return 0;
    }
    
    if (flatA.tower !== flatB.tower) {
      return flatA.tower.localeCompare(flatB.tower);
    }
    if (flatA.floor !== flatB.floor) {
      return flatA.floor - flatB.floor;
    }
    return flatA.apt - flatB.apt;
  });
  
  return sortedGroups.flatMap(([, group]) => group);
}

function createTransformFunction(isSheet2 = false) {
  return function transformRows(rows, customOrderTotals = null, customOrderItemTotals = null) {
    return rows.map(row => {
      const itemCount = parseFloat(row['Item Count']) || 0;
      const discountedPrice = parseFloat(row['Discounted Price']) || 0;
      const regularPrice = parseFloat(row['Price']) || 0;
      const rate = discountedPrice || regularPrice;
      const itemTotal = Math.round((itemCount * rate) * 100) / 100;
      const orderNum = row['Order Number'];
      
      const confirmedOrder = String(row['Confirmed Order']).toUpperCase().trim() === 'TRUE' ? 'T' : 'F';
      
      let paymentMode = '';
      let paymentStatus = 'Due';
      const originalPaymentMode = String(row['Payment Mode'] || '').trim().toLowerCase();
      const originalPaymentStatus = String(row['Payment Status'] || '').trim().toUpperCase();
      
      if (originalPaymentMode === 'phonepe' && originalPaymentStatus === 'SUCCESSFUL') {
        paymentMode = 'ONL';
        paymentStatus = 'Paid';
      }
      
      const totalAmount = customOrderTotals ? customOrderTotals.get(orderNum) : null;
      const totalItems = customOrderItemTotals ? customOrderItemTotals.get(orderNum) : null;
      
      const baseResult = {
        'Order #': row['Order Number'],
        'Flat #': row['Flat Number'],
        'Mobile No': row['Customer Mobile Number'],
        'Cnf': confirmedOrder,
        'Product Name': row['Product Name'],
        'Qty': itemCount,
        'Price': rate,
        'I Tot': itemTotal,
        'Total Items': totalItems || 0,
        'Payment Mode': paymentMode,
        'Payment Status': paymentStatus,
        'T Amt': totalAmount,
      };
      
      if (isSheet2) {
        baseResult['Catalogue Group'] = row['Catalogue Group'] || '';
        baseResult['Tax %'] = row['Tax %'] || '';
        baseResult['Tax Amount'] = row['Tax Amount'] || '';
      }
      
      return baseResult;
    });
  };
}

async function addDataToSheet(worksheet, data, addBlankRows = false, useSheet2Columns = false) {
  const columnWidths = useSheet2Columns ? COLUMN_WIDTHS_SHEET2 : COLUMN_WIDTHS;
  const columns = Object.keys(columnWidths);
  
  worksheet.columns = columns.map(col => ({
    header: col,
    key: col,
    width: columnWidths[col]
  }));
  
  const headerRow = worksheet.getRow(1);
  headerRow.height = 78;
  
  const headerStyle = {
    font: { bold: true, size: 12 },
    alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } },
    border: {
      top: { style: 'thin' },
      bottom: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' }
    }
  };
  
  columns.forEach((col, index) => {
    Object.assign(headerRow.getCell(index + 1), headerStyle);
  });
  
  let lastFlatNo = null;
  const numericColumns = new Set(['Qty', 'Price', 'I Tot', 'T Amt', 'Total Items', 'Tax %', 'Tax Amount']);
  
  data.forEach(row => {
    if (addBlankRows && lastFlatNo && lastFlatNo !== row['Flat #']) {
      const blankRow = worksheet.addRow({});
      blankRow.height = 15;
      blankRow.eachCell(cell => {
        cell.alignment = { horizontal: 'left' };
      });
    }
    
    const dataRow = worksheet.addRow(row);
    dataRow.height = 15;
    
    columns.forEach((col, index) => {
      const cell = dataRow.getCell(index + 1);
      
      cell.alignment = { horizontal: 'left' };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
      
      let font = { size: 12 };
      
      if (numericColumns.has(col)) {
        cell.numFmt = '#,##0';
        if (col === 'Qty' && parseFloat(cell.value) > 1) {
          font = { bold: true, size: 12, color: { argb: 'FFFF0000' } };
        }
      }
      
      if (col === 'Payment Status' && cell.value === 'Due') {
        font = { bold: true, size: 12, color: { argb: 'FFFF0000' } };
      }
      
      cell.font = font;
    });
    
    lastFlatNo = row['Flat #'];
  });
}

async function handleCustomerDetails(originalWorkbook, newWorkbook, filteredRows) {
  const customerSheetName = 'Cust_Data';
  const customerSheet = originalWorkbook.getWorksheet(customerSheetName);
  
  let customerData = [];
  
  if (customerSheet) {
    customerData = extractSheetData(customerSheet);
  }
  
  const existingMobileNumbers = new Set(
    customerData.map(row => row['Customer Mobile Number'] || row['Mobile Number'])
  );
  
  const newCustomers = [];
  const processedMobileNumbers = new Set();
  
  filteredRows.forEach(row => {
    const mobileNo = row['Mobile No'];
    const flatNo = row['Flat #'];
    
    if (mobileNo && !existingMobileNumbers.has(mobileNo) && 
        !processedMobileNumbers.has(mobileNo) && flatNo) {
      newCustomers.push({
        'Customer Mobile Number': mobileNo,
        'Flat Number': flatNo
      });
      processedMobileNumbers.add(mobileNo);
    }
  });
  
  if (customerData.length > 0 || newCustomers.length > 0) {
    const allCustomers = [...customerData, ...newCustomers];
    const newCustomerSheet = newWorkbook.addWorksheet(customerSheetName);
    
    newCustomerSheet.columns = [
      { header: 'Customer Mobile Number', key: 'Customer Mobile Number', width: 20 },
      { header: 'Flat Number', key: 'Flat Number', width: 15 }
    ];
    
    const headerRow = newCustomerSheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell(cell => {
      cell.font = { bold: true, size: 12 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
    
    allCustomers.forEach(customer => {
      const row = newCustomerSheet.addRow({
        'Customer Mobile Number': customer['Customer Mobile Number'] || '',
        'Flat Number': customer['Flat Number'] || ''
      });
      
      row.eachCell(cell => {
        cell.font = { size: 12 };
        cell.alignment = { horizontal: 'left' };
        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });
  }
}

// Main filterExcel function (copied and adapted from server/utils/filterExcel.js)
async function filterExcel(filePath, custDataFilePath) {
  const [workbook, custDataWorkbook] = await Promise.all([
    new ExcelJS.Workbook().xlsx.readFile(filePath),
    new ExcelJS.Workbook().xlsx.readFile(custDataFilePath)
  ]);
  
  const sheetName = 'Inquiries with order meta';
  const sheet = workbook.getWorksheet(sheetName);
  
  if (!sheet) {
    throw new Error(`Sheet '${sheetName}' not found in the uploaded file.`);
  }
  
  const data = extractSheetData(sheet);
  
  const filteredData = data.filter(row => {
    const orderStatus = String(row['Order Status'] || '').toUpperCase().trim();
    return orderStatus !== 'COMPLETED' && orderStatus !== 'REJECTED';
  });
  
  const { orderGroups, orderTotals, orderItemTotals } = processOrders(filteredData);
  
  const custDataSheet = custDataWorkbook.getWorksheet('Cust_Data');
  if (!custDataSheet) {
    throw new Error(`Sheet 'Cust_Data' not found in the customer data file.`);
  }
  
  const custLookup = new Map();
  const custHeaders = [];
  
  const custHeaderRow = custDataSheet.getRow(1);
  custHeaderRow.eachCell((cell, colNumber) => {
    custHeaders[colNumber] = cell.value;
  });
  
  custDataSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    
    let mbNo = '';
    let flatNo = '';
    
    row.eachCell((cell, colNumber) => {
      const header = custHeaders[colNumber];
      if (header === 'Mb No') mbNo = String(cell.value || '').trim();
      if (header === 'Flat No') flatNo = String(cell.value || '').trim();
    });
    
    if (mbNo) custLookup.set(mbNo, flatNo);
  });
  
  const flaggedOrders = [];
  const validOrders = [];
  
  filteredData.forEach(row => {
    const mobileNo = String(row['Customer Mobile Number'] || '').trim();
    const shippingAddress = String(row['Flat Number'] || '').trim();
    const lookupFlatNo = custLookup.get(mobileNo);
    
    if (lookupFlatNo) {
      if (!shippingAddress) {
        validOrders.push(row);
        return;
      }
      
      const addressNumber = extractNumberFromAddress(shippingAddress);
      const flatNumber = extractNumberFromAddress(lookupFlatNo);
      
      if (addressNumber === flatNumber) {
        validOrders.push(row);
      } else {
        flaggedOrders.push(row);
      }
    } else {
      validOrders.push(row);
    }
  });
  
  let flaggedOrderTotals = new Map();
  let flaggedOrderItemTotals = new Map();
  
  if (flaggedOrders.length > 0) {
    const flaggedProcessed = processOrders(flaggedOrders);
    flaggedOrderTotals = flaggedProcessed.orderTotals;
    flaggedOrderItemTotals = flaggedProcessed.orderItemTotals;
  }
  
  const mainOrders = [];
  const newNumOrders = [];
  
  validOrders.forEach(row => {
    const mobileNo = String(row['Customer Mobile Number'] || '').trim();
    if (custLookup.has(mobileNo)) {
      row['Flat Number'] = custLookup.get(mobileNo);
      mainOrders.push(row);
    } else {
      newNumOrders.push(row);
    }
  });
  
  const sortedMainOrders = groupAndSort(mainOrders, true);
  const sortedNewNumOrders = groupAndSort(newNumOrders);
  
  const transformRows = createTransformFunction(false);
  const transformRowsSheet2 = createTransformFunction(true);
  
  const transformedMainOrders = transformRows(sortedMainOrders, orderTotals, orderItemTotals);
  const transformedNewNumOrders = transformRows(sortedNewNumOrders, orderTotals, orderItemTotals);
  const allTransformedOrders = [...transformedMainOrders, ...transformedNewNumOrders];
  
  const newWorkbook = new ExcelJS.Workbook();
  
  const sheetPromises = [];
  
  if (flaggedOrders.length > 0) {
    const flaggedSheet = newWorkbook.addWorksheet('Flagged_Add');
    const transformedFlaggedOrders = transformRows(flaggedOrders, flaggedOrderTotals, flaggedOrderItemTotals);
    sheetPromises.push(addDataToSheet(flaggedSheet, transformedFlaggedOrders));
  }
  
  const mainSheet = newWorkbook.addWorksheet('Sheet1');
  sheetPromises.push(addDataToSheet(mainSheet, transformedMainOrders));
  
  const allValidOrders = [...sortedMainOrders, ...sortedNewNumOrders];
  const transformedSheet2Orders = transformRowsSheet2(allValidOrders, orderTotals, orderItemTotals);
  const sheet2 = newWorkbook.addWorksheet('Sheet2');
  sheetPromises.push(addDataToSheet(sheet2, transformedSheet2Orders, false, true));
  
  if (transformedNewNumOrders.length > 0) {
    const newNumSheet = newWorkbook.addWorksheet('New_Num');
    sheetPromises.push(addDataToSheet(newNumSheet, transformedNewNumOrders));
  }
  
  const towerData = new Map();
  const towers = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'M', 'N', 'P'];
  
  allTransformedOrders.forEach(row => {
    const flatNo = String(row['Flat #'] || '').toUpperCase();
    for (const tower of towers) {
      if (flatNo.startsWith(tower)) {
        if (!towerData.has(tower)) {
          towerData.set(tower, []);
        }
        towerData.get(tower).push(row);
        break;
      }
    }
  });
  
  for (const [tower, data] of towerData) {
    const towerSheet = newWorkbook.addWorksheet(`Tower ${tower}`);
    sheetPromises.push(addDataToSheet(towerSheet, data, true));
  }
  
  await Promise.all(sheetPromises);
  
  await handleCustomerDetails(workbook, newWorkbook, allTransformedOrders);
  
  const buffer = await newWorkbook.xlsx.writeBuffer();
  return buffer;
}

export default async function handler(req, res) {
  // Enable CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  let tempFiles = [];

  try {
    // Get the request body
    const contentType = req.headers['content-type'] || '';
    console.log('Content-Type:', contentType);
    
    const boundary = contentType.split('boundary=')[1];
    
    if (!boundary) {
      console.error('No boundary found in content-type:', contentType);
      res.status(400).json({ error: 'No boundary found in multipart data' });
      return;
    }
    
    console.log('Boundary:', boundary);

    // Read the request body as buffer
    const chunks = [];
    
    for await (const chunk of req) {
      chunks.push(chunk);
    }
    
    const bodyBuffer = Buffer.concat(chunks);
    console.log('Body buffer length:', bodyBuffer.length);

    // Parse multipart data
    const files = parseMultipartData(bodyBuffer, boundary);
    console.log('Parsed files:', Object.keys(files));
    
    const mainFile = files.file;
    const custDataFile = files.custData;

    if (!mainFile || !custDataFile) {
      console.error('Files missing - file:', !!mainFile, 'custData:', !!custDataFile);
      res.status(400).json({ 
        error: 'Both files must be uploaded',
        received: Object.keys(files)
      });
      return;
    }
    
    console.log('File sizes - main:', mainFile.buffer.length, 'custData:', custDataFile.buffer.length);

    // Create unique temporary file paths
    const timestamp = Date.now();
    const randomId = Math.random().toString(36).substr(2, 9);
    const mainFilePath = path.join('/tmp', `main_${timestamp}_${randomId}.xlsx`);
    const custDataFilePath = path.join('/tmp', `cust_${timestamp}_${randomId}.xlsx`);

    // Write uploaded files to temp paths
    fs.writeFileSync(mainFilePath, mainFile.buffer);
    fs.writeFileSync(custDataFilePath, custDataFile.buffer);
    
    tempFiles = [mainFilePath, custDataFilePath];

    // Process the Excel files
    const filteredBuffer = await filterExcel(mainFilePath, custDataFilePath);

    // Set response headers for file download
    res.setHeader('Content-Disposition', 'attachment; filename="filtered.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Length', filteredBuffer.length);
    
    res.status(200).send(filteredBuffer);

  } catch (error) {
    console.error('Error processing files:', error);
    res.status(500).json({ 
      error: 'Error processing files: ' + error.message 
    });
  } finally {
    // Clean up temporary files
    tempFiles.forEach(filePath => {
      if (filePath && fs.existsSync(filePath)) {
        try {
          fs.unlinkSync(filePath);
        } catch (err) {
          console.error('Error deleting temp file:', err);
        }
      }
    });
  }
}
