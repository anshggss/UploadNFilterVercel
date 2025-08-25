import ExcelJS from 'exceljs';

// Column mappings
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

// Adjusted column widths to prevent ## displaying
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

// Column widths for Sheet2 (includes additional columns)
const COLUMN_WIDTHS_SHEET2 = {
  ...COLUMN_WIDTHS,
  'Catalogue Group': 20,
  'Tax %': 8,
  'Tax Amount': 10
};

// Cache for normalized strings to avoid repeated processing
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

// Optimized flat number parser with caching
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

// Optimized data extraction with single pass
function extractSheetData(sheet) {
  const data = [];
  const headers = [];
  
  // Get headers once
  const headerRow = sheet.getRow(1);
  headerRow.eachCell((cell, colNumber) => {
    headers[colNumber] = cell.value;
  });
  
  // Process data rows
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // Skip header
    
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

// Optimized order processing with single pass calculations
function processOrders(data) {
  const orderGroups = new Map();
  const orderTotals = new Map();
  const orderItemTotals = new Map();
  
  // Single pass to group orders and calculate totals
  data.forEach(row => {
    const orderNum = row['Order Number'];
    
    if (!orderGroups.has(orderNum)) {
      orderGroups.set(orderNum, []);
      orderTotals.set(orderNum, 0);
      orderItemTotals.set(orderNum, 0);
    }
    
    orderGroups.get(orderNum).push(row);
    
    // Calculate totals incrementally
    const count = parseFloat(row['Item Count']) || 0;
    const discountedPrice = parseFloat(row['Discounted Price']) || 0;
    const regularPrice = parseFloat(row['Price']) || 0;
    const price = discountedPrice || regularPrice;
    
    orderTotals.set(orderNum, orderTotals.get(orderNum) + (count * price));
    orderItemTotals.set(orderNum, orderItemTotals.get(orderNum) + count);
  });
  
  // Round totals
  for (const [orderNum, total] of orderTotals) {
    orderTotals.set(orderNum, Math.round(total * 100) / 100);
  }
  
  return { orderGroups, orderTotals, orderItemTotals };
}

// Optimized sorting function
function groupAndSort(orders, moveEmptyToBottom = false) {
  // Group by Order Number using Map for better performance
  const groups = new Map();
  
  orders.forEach(row => {
    const orderNum = row['Order Number'];
    if (!groups.has(orderNum)) {
      groups.set(orderNum, []);
    }
    groups.get(orderNum).push(row);
  });
  
  // Convert to array and sort
  const sortedGroups = Array.from(groups.entries()).sort((a, b) => {
    const flatA = parseFlatNumber(a[1][0]['Flat Number']);
    const flatB = parseFlatNumber(b[1][0]['Flat Number']);
    
    if (moveEmptyToBottom) {
      if (flatA.isEmpty && !flatB.isEmpty) return 1;
      if (!flatA.isEmpty && flatB.isEmpty) return -1;
      if (flatA.isEmpty && flatB.isEmpty) return 0;
    }
    
    // Multi-level comparison
    if (flatA.tower !== flatB.tower) {
      return flatA.tower.localeCompare(flatB.tower);
    }
    if (flatA.floor !== flatB.floor) {
      return flatA.floor - flatB.floor;
    }
    return flatA.apt - flatB.apt;
  });
  
  // Flatten efficiently
  return sortedGroups.flatMap(([, group]) => group);
}

// Optimized transformation with pre-computed values
function createTransformFunction(isSheet2 = false) {
  return function transformRows(rows, customOrderTotals = null, customOrderItemTotals = null) {
    return rows.map(row => {
      const itemCount = parseFloat(row['Item Count']) || 0;
      const discountedPrice = parseFloat(row['Discounted Price']) || 0;
      const regularPrice = parseFloat(row['Price']) || 0;
      const rate = discountedPrice || regularPrice;
      const itemTotal = Math.round((itemCount * rate) * 100) / 100;
      const orderNum = row['Order Number'];
      
      // Optimized boolean conversion
      const confirmedOrder = String(row['Confirmed Order']).toUpperCase().trim() === 'TRUE' ? 'T' : 'F';
      
      // Optimized payment processing
      let paymentMode = '';
      let paymentStatus = 'Due';
      const originalPaymentMode = String(row['Payment Mode'] || '').trim().toLowerCase();
      const originalPaymentStatus = String(row['Payment Status'] || '').trim().toUpperCase();
      
      if (originalPaymentMode === 'phonepe' && originalPaymentStatus === 'SUCCESSFUL') {
        paymentMode = 'ONL';
        paymentStatus = 'Paid';
      }
      
      // Use Maps for O(1) lookup
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

export async function filterExcel(filePath, custDataFilePath) {
  // Read workbooks in parallel
  const [workbook, custDataWorkbook] = await Promise.all([
    new ExcelJS.Workbook().xlsx.readFile(filePath),
    new ExcelJS.Workbook().xlsx.readFile(custDataFilePath)
  ]);
  
  const sheetName = 'Inquiries with order meta';
  const sheet = workbook.getWorksheet(sheetName);
  
  if (!sheet) {
    throw new Error(`Sheet '${sheetName}' not found in the uploaded file.`);
  }
  
  // Extract data efficiently
  const data = extractSheetData(sheet);
  
  // Filter orders efficiently
  const filteredData = data.filter(row => {
    const orderStatus = String(row['Order Status'] || '').toUpperCase().trim();
    return orderStatus !== 'COMPLETED' && orderStatus !== 'REJECTED';
  });
  
  // Process orders with single pass
  const { orderGroups, orderTotals, orderItemTotals } = processOrders(filteredData);
  
  // Build customer lookup efficiently
  const custDataSheet = custDataWorkbook.getWorksheet('Cust_Data');
  if (!custDataSheet) {
    throw new Error(`Sheet 'Cust_Data' not found in the customer data file.`);
  }
  
  const custLookup = new Map();
  const custHeaders = [];
  
  // Get headers once
  const custHeaderRow = custDataSheet.getRow(1);
  custHeaderRow.eachCell((cell, colNumber) => {
    custHeaders[colNumber] = cell.value;
  });
  
  // Build lookup map
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
  
  // Address validation with optimized processing
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
  
  // Process flagged orders if any exist
  let flaggedOrderTotals = new Map();
  let flaggedOrderItemTotals = new Map();
  
  if (flaggedOrders.length > 0) {
    const flaggedProcessed = processOrders(flaggedOrders);
    flaggedOrderTotals = flaggedProcessed.orderTotals;
    flaggedOrderItemTotals = flaggedProcessed.orderItemTotals;
  }
  
  // Separate main and new orders efficiently
  const mainOrders = [];
  const newNumOrders = [];
  
  validOrders.forEach(row => {
    const mobileNo = String(row['Customer Mobile Number'] || '').trim();
    if (custLookup.has(mobileNo)) {
      // Update flat number for main orders
      row['Flat Number'] = custLookup.get(mobileNo);
      mainOrders.push(row);
    } else {
      newNumOrders.push(row);
    }
  });
  
  // Sort orders
  const sortedMainOrders = groupAndSort(mainOrders, true);
  const sortedNewNumOrders = groupAndSort(newNumOrders);
  
  // Create transform functions
  const transformRows = createTransformFunction(false);
  const transformRowsSheet2 = createTransformFunction(true);
  
  // Transform data
  const transformedMainOrders = transformRows(sortedMainOrders, orderTotals, orderItemTotals);
  const transformedNewNumOrders = transformRows(sortedNewNumOrders, orderTotals, orderItemTotals);
  const allTransformedOrders = [...transformedMainOrders, ...transformedNewNumOrders];
  
  // Create workbook and sheets
  const newWorkbook = new ExcelJS.Workbook();
  
  // Create sheets in parallel where possible
  const sheetPromises = [];
  
  // Flagged_Add sheet
  if (flaggedOrders.length > 0) {
    const flaggedSheet = newWorkbook.addWorksheet('Flagged_Add');
    const transformedFlaggedOrders = transformRows(flaggedOrders, flaggedOrderTotals, flaggedOrderItemTotals);
    sheetPromises.push(addDataToSheet(flaggedSheet, transformedFlaggedOrders));
  }
  
  // Main sheet
  const mainSheet = newWorkbook.addWorksheet('Sheet1');
  sheetPromises.push(addDataToSheet(mainSheet, transformedMainOrders));
  
  // Sheet2
  const allValidOrders = [...sortedMainOrders, ...sortedNewNumOrders];
  const transformedSheet2Orders = transformRowsSheet2(allValidOrders, orderTotals, orderItemTotals);
  const sheet2 = newWorkbook.addWorksheet('Sheet2');
  sheetPromises.push(addDataToSheet(sheet2, transformedSheet2Orders, false, true));
  
  // New_Num sheet
  if (transformedNewNumOrders.length > 0) {
    const newNumSheet = newWorkbook.addWorksheet('New_Num');
    sheetPromises.push(addDataToSheet(newNumSheet, transformedNewNumOrders));
  }
  
  // Tower sheets - group data first to avoid redundant filtering
  const towerData = new Map();
  const towers = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'M', 'N', 'P'];
  
  // Group by tower in single pass
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
  
  // Create tower sheets
  for (const [tower, data] of towerData) {
    const towerSheet = newWorkbook.addWorksheet(`Tower ${tower}`);
    sheetPromises.push(addDataToSheet(towerSheet, data, true));
  }
  
  // Wait for all sheet operations to complete
  await Promise.all(sheetPromises);
  
  // Handle customer details
  await handleCustomerDetails(workbook, newWorkbook, allTransformedOrders);
  
  // Write to buffer
  const buffer = await newWorkbook.xlsx.writeBuffer();
  return buffer;
}

// Optimized sheet data addition with batch operations
async function addDataToSheet(worksheet, data, addBlankRows = false, useSheet2Columns = false) {
  const columnWidths = useSheet2Columns ? COLUMN_WIDTHS_SHEET2 : COLUMN_WIDTHS;
  const columns = Object.keys(columnWidths);
  
  // Set up columns
  worksheet.columns = columns.map(col => ({
    header: col,
    key: col,
    width: columnWidths[col]
  }));
  
  // Style header row
  const headerRow = worksheet.getRow(1);
  headerRow.height = 78;
  
  // Batch header formatting
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
  
  // Add data rows efficiently
  let lastFlatNo = null;
  const numericColumns = new Set(['Qty', 'Price', 'I Tot', 'T Amt', 'Total Items', 'Tax %', 'Tax Amount']);
  
  data.forEach(row => {
    // Add blank row when flat number changes
    if (addBlankRows && lastFlatNo && lastFlatNo !== row['Flat #']) {
      const blankRow = worksheet.addRow({});
      blankRow.height = 15;
      blankRow.eachCell(cell => {
        cell.alignment = { horizontal: 'left' };
      });
    }
    
    const dataRow = worksheet.addRow(row);
    dataRow.height = 15;
    
    // Batch cell formatting
    columns.forEach((col, index) => {
      const cell = dataRow.getCell(index + 1);
      
      // Base formatting
      cell.alignment = { horizontal: 'left' };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
      
      let font = { size: 12 };
      
      // Conditional formatting
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

// Optimized customer details handling
async function handleCustomerDetails(originalWorkbook, newWorkbook, filteredRows) {
  const customerSheetName = 'Cust_Data';
  const customerSheet = originalWorkbook.getWorksheet(customerSheetName);
  
  let customerData = [];
  
  if (customerSheet) {
    customerData = extractSheetData(customerSheet);
  }
  
  // Use Set for O(1) lookup
  const existingMobileNumbers = new Set(
    customerData.map(row => row['Customer Mobile Number'] || row['Mobile Number'])
  );
  
  // Find new customers efficiently
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
    
    // Set up columns
    newCustomerSheet.columns = [
      { header: 'Customer Mobile Number', key: 'Customer Mobile Number', width: 20 },
      { header: 'Flat Number', key: 'Flat Number', width: 15 }
    ];
    
    // Style header
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
    
    // Add customer data
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