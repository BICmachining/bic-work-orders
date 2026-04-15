// BIC Machine Shop - Work Orders Apps Script
// Paste this entire file into your Google Apps Script editor
// Then deploy as a Web App (instructions in README.md)
// IMPORTANT: After any changes, redeploy as a NEW version and update the URL in Settings on all devices.

const SHEET_ID = '16-aMTlmhRaWvdZzBifXMJEVo5IA60TWcHC8rhHcLGQw';
const PO_SHEET_NAME = 'POs';
const LINE_ITEMS_SHEET_NAME = 'Line Items';
const TOOLS_SHEET_NAME = 'Tools';
const INSERTS_SHEET_NAME = 'Inserts';
const HOLDERS_SHEET_NAME = 'Holders';
const ACCESSORIES_SHEET_NAME = 'Accessories';
const INSERT_TOOL_LINKS_SHEET_NAME = 'Insert-Tool Links';
const TOOL_PART_LINKS_SHEET_NAME = 'Tool-Part Links';
const TOOL_MACHINE_INV_SHEET_NAME = 'Tool-Machine Inventory';
const MACHINES_SHEET_NAME = 'Machines';

// ─── Router ────────────────────────────────────────────────────────────────

function doGet(e) {
  try {
    const action = e.parameter.action;

    if (action === 'submitPO') {
      const po = JSON.parse(decodeURIComponent(e.parameter.po));
      const lineItems = JSON.parse(decodeURIComponent(e.parameter.lineItems));
      return submitPO(po, lineItems);
    } else if (action === 'updateStatus') {
      const poNumber = e.parameter.poNumber;
      const line = e.parameter.line;
      const status = e.parameter.status;
      const notes = e.parameter.notes || '';
      return updateStatus(poNumber, line, status, notes);
    } else if (action === 'getDashboardData') {
      return getDashboardData();

    // ── Tooling actions ──────────────────────────────────────────────────
    } else if (action === 'getToolingData') {
      return getToolingData();
    } else if (action === 'addTool') {
      return addTool(e.parameter);
    } else if (action === 'updateTool') {
      return updateTool(e.parameter);
    } else if (action === 'addInsert') {
      return addInsert(e.parameter);
    } else if (action === 'updateInsert') {
      return updateInsert(e.parameter);
    } else if (action === 'addHolder') {
      return addHolder(e.parameter);
    } else if (action === 'updateHolder') {
      return updateHolder(e.parameter);
    } else if (action === 'addAccessory') {
      return addAccessory(e.parameter);
    } else if (action === 'addInsertToolLink') {
      return addInsertToolLink(e.parameter);
    } else if (action === 'removeInsertToolLink') {
      return removeInsertToolLink(e.parameter);
    } else if (action === 'addToolPartLink') {
      return addToolPartLink(e.parameter);
    } else if (action === 'removeToolPartLink') {
      return removeToolPartLink(e.parameter);
    } else if (action === 'updateToolMachineInventory') {
      return updateToolMachineInventory(e.parameter);
    } else if (action === 'updateToolQty') {
      return updateToolQty(e.parameter);
    } else if (action === 'updateInsertQty') {
      return updateInsertQty(e.parameter);
    } else if (action === 'addMachine') {
      return addMachine(e.parameter);
    } else if (action === 'getReorderData') {
      return getReorderData();
    } else if (action === 'deletePO') {
      return deletePO(e.parameter);
    } else if (action === 'updateTimes') {
      return updateTimes(e.parameter);

    } else {
      return response({ success: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return response({ success: false, error: err.toString() });
  }
}

// ─── Helpers ───────────────────────────────────────────────────────────────

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return (val.getMonth() + 1) + '/' + val.getDate() + '/' + String(val.getFullYear()).slice(-2);
  }
  return String(val);
}

function generateId(prefix, sheet) {
  const lastRow = sheet.getLastRow();
  // Scan all existing IDs in column A to find the true max — handles manual edits,
  // deletions, and gaps in the sequence so we never collide with an existing ID.
  if (lastRow <= 1) return prefix + '-001';
  const idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const pattern = new RegExp('^' + prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '-(\\d+)$');
  let maxNum = 0;
  idValues.forEach(function(row) {
    const m = String(row[0]).match(pattern);
    if (m) {
      const n = parseInt(m[1], 10);
      if (n > maxNum) maxNum = n;
    }
  });
  return prefix + '-' + String(maxNum + 1).padStart(3, '0');
}

function response(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Step 1 helper: run this ONCE to create all tooling sheet tabs ──────────
// In Apps Script editor: run initializeToolingSheets() manually from the menu.

function initializeToolingSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const sheetsToCreate = [
    {
      name: TOOLS_SHEET_NAME,
      headers: ['Tool ID','Type','Tooling Type','Description','Vendor','Product #','Product URL',
                'Diameter','Diameter Unit','Flutes','OAL','OAL Unit','Corner Radius','Tip Angle',
                'Taper Angle','Thread Pitch','Body Material','Stock Qty','Min Qty','Last Price',
                'Last Order Date','Fusion GUID','Notes']
    },
    {
      name: INSERTS_SHEET_NAME,
      headers: ['Insert ID','Description','Vendor','Product #','Product URL','Insert Geometry',
                'Grade','Material Application','Nose Radius','IC Size','Qty On Hand','Min Qty',
                'Last Price','Last Order Date','Notes']
    },
    {
      name: HOLDERS_SHEET_NAME,
      headers: ['Holder ID','Description','Vendor','Product #','Product URL','Type','Taper',
                'Bore/Capacity','Qty On Hand','Min Qty','Last Price','Last Order Date','Notes']
    },
    {
      name: ACCESSORIES_SHEET_NAME,
      headers: ['Accessory ID','Tool ID','Type','Description','Vendor','Product #',
                'Thread Size','Drive Type','Notes']
    },
    {
      name: INSERT_TOOL_LINKS_SHEET_NAME,
      headers: ['Link ID','Tool ID','Insert ID','Notes']
    },
    {
      name: TOOL_PART_LINKS_SHEET_NAME,
      headers: ['Link ID','Part Number','Tool ID','Fusion T#','Source','Fusion File','Notes']
    },
    {
      name: TOOL_MACHINE_INV_SHEET_NAME,
      headers: ['Record ID','Tool ID','Machine ID','Qty','In Magazine (Y/N)']
    },
    {
      name: MACHINES_SHEET_NAME,
      headers: ['Machine ID','Name','Type','Notes']
    }
  ];

  sheetsToCreate.forEach(function(def) {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
    }
    // Write headers only if row 1 is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(def.headers);
      sheet.getRange(1, 1, 1, def.headers.length).setFontWeight('bold');
    }
  });

  // Seed Machines tab
  const machinesSheet = ss.getSheetByName(MACHINES_SHEET_NAME);
  if (machinesSheet.getLastRow() <= 1) {
    machinesSheet.appendRow(['MCH-001', 'Haas VF-4SS',       'Vertical Machining Center', '']);
    machinesSheet.appendRow(['MCH-002', 'Haas ST-25Y',       'CNC Lathe',                 '']);
    machinesSheet.appendRow(['MCH-003', 'Dah Lih Bed Mill',  'Manual Mill',               '']);
  }

  Logger.log('Tooling sheets initialized successfully.');
}

// ─── Existing PO actions ───────────────────────────────────────────────────

function submitPO(po, lineItems) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const poSheet = ss.getSheetByName(PO_SHEET_NAME);
  const lineSheet = ss.getSheetByName(LINE_ITEMS_SHEET_NAME);

  const dateAdded = new Date().toLocaleDateString('en-US');

  // Check if PO already exists
  const poData = poSheet.getDataRange().getValues();
  for (let i = 1; i < poData.length; i++) {
    if (String(poData[i][0]) === String(po.poNumber)) {
      return response({ success: false, error: 'PO ' + po.poNumber + ' already exists in the sheet.' });
    }
  }

  // Write PO header row
  poSheet.appendRow([
    po.poNumber,
    po.poDate,
    po.orderedBy,
    po.orderDescription,
    po.handwrittenNotes || '',
    dateAdded
  ]);

  // Write line item rows
  lineItems.forEach(item => {
    lineSheet.appendRow([
      po.poNumber,
      item.line,
      item.partNumber,
      item.qty,
      item.description,
      item.materialPartNumber || '',
      item.materialDescription || '',
      item.materialUsed || '',
      item.materialPerPart || '',
      item.dueDate,
      'Not Started',
      item.notes || ''
    ]);
  });

  return response({ success: true, message: 'PO ' + po.poNumber + ' saved successfully.' });
}

function updateStatus(poNumber, line, status, notes) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const lineSheet = ss.getSheetByName(LINE_ITEMS_SHEET_NAME);
  const values = lineSheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(poNumber) && String(values[i][1]) === String(line)) {
      lineSheet.getRange(i + 1, 11).setValue(status);                         // col K: Status
      if (notes !== undefined && notes !== '') {
        lineSheet.getRange(i + 1, 12).setValue(notes);                         // col L: Notes
      }
      lineSheet.getRange(i + 1, 13).setValue(new Date().toLocaleDateString('en-US')); // col M: Status Updated
      return response({ success: true });
    }
  }
  return response({ success: false, error: 'Line item not found.' });
}

function updateTimes(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const lineSheet = ss.getSheetByName(LINE_ITEMS_SHEET_NAME);
  const values = lineSheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(p.poNumber) && String(values[i][1]) === String(p.line)) {
      if (p.cyclePerPart !== undefined && p.cyclePerPart !== '') {
        lineSheet.getRange(i + 1, 14).setValue(Number(p.cyclePerPart)); // col N: Cycle Per Part
      }
      if (p.totalJobTime !== undefined && p.totalJobTime !== '') {
        lineSheet.getRange(i + 1, 15).setValue(Number(p.totalJobTime)); // col O: Total Job Time
      }
      if (p.qtyMade !== undefined && p.qtyMade !== '') {
        lineSheet.getRange(i + 1, 16).setValue(Number(p.qtyMade)); // col P: Qty Made
      }
      return response({ success: true });
    }
  }
  return response({ success: false, error: 'Line item not found.' });
}

function deletePO(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const poNumber = String(p.poNumber || '').trim();
  if (!poNumber) return response({ success: false, error: 'No PO number provided.' });

  const poSheet   = ss.getSheetByName(PO_SHEET_NAME);
  const lineSheet = ss.getSheetByName(LINE_ITEMS_SHEET_NAME);

  // Delete the PO header row (scan from bottom so row index stays valid)
  const poValues = poSheet.getDataRange().getValues();
  for (let i = poValues.length - 1; i >= 1; i--) {
    if (String(poValues[i][0]).trim() === poNumber) {
      poSheet.deleteRow(i + 1);
      break;
    }
  }

  // Delete all line items for this PO (scan from bottom)
  const lineValues = lineSheet.getDataRange().getValues();
  for (let i = lineValues.length - 1; i >= 1; i--) {
    if (String(lineValues[i][0]).trim() === poNumber) {
      lineSheet.deleteRow(i + 1);
    }
  }

  return response({ success: true, deleted: poNumber });
}

function getDashboardData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const poSheet = ss.getSheetByName(PO_SHEET_NAME);
  const lineSheet = ss.getSheetByName(LINE_ITEMS_SHEET_NAME);

  const poRows = poSheet.getDataRange().getValues();
  const lineRows = lineSheet.getDataRange().getValues();

  const pos = poRows.slice(1).map(r => ({
    poNumber: r[0],
    poDate: formatDate(r[1]),
    orderedBy: r[2],
    orderDescription: r[3],
    handwrittenNotes: r[4],
    dateAdded: formatDate(r[5])
  }));

  const lineItems = lineRows.slice(1).map(r => ({
    poNumber:           r[0],
    line:               r[1],
    partNumber:         r[2],
    qty:                r[3],
    description:        r[4],
    materialPartNumber: r[5],
    materialDescription:r[6],
    materialUsed:       r[7],
    materialPerPart:    r[8],
    dueDate:            formatDate(r[9]),
    status:             r[10],
    notes:              r[11],
    statusUpdated:      formatDate(r[12]),  // col M
    cyclePerPart:       r[13] !== undefined && r[13] !== '' ? r[13] : '',  // col N
    totalJobTime:       r[14] !== undefined && r[14] !== '' ? r[14] : '',  // col O
    qtyMade:            r[15] !== undefined && r[15] !== '' ? r[15] : ''   // col P
  }));

  return response({ success: true, pos: pos, lineItems: lineItems });
}

// ─── Tooling: read ─────────────────────────────────────────────────────────

function getToolingData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  function sheetToObjects(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];
    const headers = rows[0];
    return rows.slice(1).map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
  }

  const tools    = sheetToObjects(TOOLS_SHEET_NAME).map(r => ({
    toolId:         r['Tool ID'],
    type:           r['Type'],
    toolingType:    r['Tooling Type'],
    description:    r['Description'],
    vendor:         r['Vendor'],
    productNumber:  r['Product #'],
    productUrl:     r['Product URL'],
    diameter:       r['Diameter'],
    diameterUnit:   r['Diameter Unit'],
    flutes:         r['Flutes'],
    oal:            r['OAL'],
    oalUnit:        r['OAL Unit'],
    cornerRadius:   r['Corner Radius'],
    tipAngle:       r['Tip Angle'],
    taperAngle:     r['Taper Angle'],
    threadPitch:    r['Thread Pitch'],
    bodyMaterial:   r['Body Material'],
    stockQty:       r['Stock Qty'],
    minQty:         r['Min Qty'],
    lastPrice:      r['Last Price'],
    lastOrderDate:  formatDate(r['Last Order Date']),
    fusionGuid:     r['Fusion GUID'],
    notes:          r['Notes']
  }));

  const inserts  = sheetToObjects(INSERTS_SHEET_NAME).map(r => ({
    insertId:           r['Insert ID'],
    description:        r['Description'],
    vendor:             r['Vendor'],
    productNumber:      r['Product #'],
    productUrl:         r['Product URL'],
    insertGeometry:     r['Insert Geometry'],
    grade:              r['Grade'],
    materialApplication:r['Material Application'],
    noseRadius:         r['Nose Radius'],
    icSize:             r['IC Size'],
    qtyOnHand:          r['Qty On Hand'],
    minQty:             r['Min Qty'],
    lastPrice:          r['Last Price'],
    lastOrderDate:      formatDate(r['Last Order Date']),
    notes:              r['Notes']
  }));

  const holders  = sheetToObjects(HOLDERS_SHEET_NAME).map(r => ({
    holderId:       r['Holder ID'],
    description:    r['Description'],
    vendor:         r['Vendor'],
    productNumber:  r['Product #'],
    productUrl:     r['Product URL'],
    type:           r['Type'],
    taper:          r['Taper'],
    boreCapacity:   r['Bore/Capacity'],
    qtyOnHand:      r['Qty On Hand'],
    minQty:         r['Min Qty'],
    lastPrice:      r['Last Price'],
    lastOrderDate:  formatDate(r['Last Order Date']),
    notes:          r['Notes']
  }));

  const accessories = sheetToObjects(ACCESSORIES_SHEET_NAME).map(r => ({
    accessoryId:  r['Accessory ID'],
    toolId:       r['Tool ID'],
    type:         r['Type'],
    description:  r['Description'],
    vendor:       r['Vendor'],
    productNumber:r['Product #'],
    threadSize:   r['Thread Size'],
    driveType:    r['Drive Type'],
    notes:        r['Notes']
  }));

  const insertToolLinks = sheetToObjects(INSERT_TOOL_LINKS_SHEET_NAME).map(r => ({
    linkId:   r['Link ID'],
    toolId:   r['Tool ID'],
    insertId: r['Insert ID'],
    notes:    r['Notes']
  }));

  const toolPartLinks = sheetToObjects(TOOL_PART_LINKS_SHEET_NAME).map(r => ({
    linkId:     r['Link ID'],
    partNumber: r['Part Number'],
    toolId:     r['Tool ID'],
    fusionTNum: r['Fusion T#'],
    source:     r['Source'],
    fusionFile: r['Fusion File'],
    notes:      r['Notes']
  }));

  const toolMachineInventory = sheetToObjects(TOOL_MACHINE_INV_SHEET_NAME).map(r => ({
    recordId:   r['Record ID'],
    toolId:     r['Tool ID'],
    machineId:  r['Machine ID'],
    qty:        r['Qty'],
    inMagazine: r['In Magazine (Y/N)']
  }));

  const machines = sheetToObjects(MACHINES_SHEET_NAME).map(r => ({
    machineId: r['Machine ID'],
    name:      r['Name'],
    type:      r['Type'],
    notes:     r['Notes']
  }));

  return response({
    success: true,
    tools,
    inserts,
    holders,
    accessories,
    insertToolLinks,
    toolPartLinks,
    toolMachineInventory,
    machines
  });
}

// ─── Tooling: Tools ────────────────────────────────────────────────────────

function addTool(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TOOLS_SHEET_NAME);
  // Duplicate check: reject if a tool with this product number already exists
  if (p.productNumber && String(p.productNumber).trim() !== '') {
    const existing = sheet.getDataRange().getValues();
    for (var di = 1; di < existing.length; di++) {
      if (String(existing[di][5]).trim().toLowerCase() === String(p.productNumber).trim().toLowerCase()) {
        return response({ success: false, error: 'Duplicate: product # ' + p.productNumber + ' already exists as ' + existing[di][0] + ' (' + existing[di][3] + ')' });
      }
    }
  }
  const toolId = generateId('TL', sheet);
  sheet.appendRow([
    toolId,
    p.type || '',
    p.toolingType || '',
    p.description || '',
    p.vendor || '',
    p.productNumber || '',
    p.productUrl || '',
    p.diameter || '',
    p.diameterUnit || '',
    p.flutes || '',
    p.oal || '',
    p.oalUnit || '',
    p.cornerRadius || '',
    p.tipAngle || '',
    p.taperAngle || '',
    p.threadPitch || '',
    p.bodyMaterial || '',
    p.stockQty || 0,
    p.minQty || 0,
    p.lastPrice || '',
    p.lastOrderDate || '',
    p.fusionGuid || '',
    p.notes || ''
  ]);
  return response({ success: true, toolId: toolId });
}

function updateTool(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TOOLS_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(p.toolId)) {
      const row = i + 1;
      const updates = {
        2:  p.type,         3:  p.toolingType,  4:  p.description,
        5:  p.vendor,       6:  p.productNumber, 7:  p.productUrl,
        8:  p.diameter,     9:  p.diameterUnit,  10: p.flutes,
        11: p.oal,          12: p.oalUnit,       13: p.cornerRadius,
        14: p.tipAngle,     15: p.taperAngle,    16: p.threadPitch,
        17: p.bodyMaterial, 18: p.stockQty,      19: p.minQty,
        20: p.lastPrice,    21: p.lastOrderDate, 22: p.fusionGuid,
        23: p.notes
      };
      Object.entries(updates).forEach(([col, val]) => {
        if (val !== undefined) sheet.getRange(row, Number(col)).setValue(val);
      });
      return response({ success: true });
    }
  }
  return response({ success: false, error: 'Tool not found: ' + p.toolId });
}

function updateToolQty(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TOOLS_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(p.toolId)) {
      sheet.getRange(i + 1, 18).setValue(Number(p.qty)); // col R = Stock Qty
      return response({ success: true });
    }
  }
  return response({ success: false, error: 'Tool not found: ' + p.toolId });
}

// ─── Tooling: Inserts ──────────────────────────────────────────────────────

function addInsert(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(INSERTS_SHEET_NAME);
  // Duplicate check by product number
  if (p.productNumber && String(p.productNumber).trim() !== '') {
    const existing = sheet.getDataRange().getValues();
    for (var di = 1; di < existing.length; di++) {
      if (String(existing[di][3]).trim().toLowerCase() === String(p.productNumber).trim().toLowerCase()) {
        return response({ success: false, error: 'Duplicate: insert product # ' + p.productNumber + ' already exists as ' + existing[di][0] });
      }
    }
  }
  const insertId = generateId('IN', sheet);
  sheet.appendRow([
    insertId,
    p.description || '',
    p.vendor || '',
    p.productNumber || '',
    p.productUrl || '',
    p.insertGeometry || '',
    p.grade || '',
    p.materialApplication || '',
    p.noseRadius || '',
    p.icSize || '',
    p.qtyOnHand || 0,
    p.minQty || 0,
    p.lastPrice || '',
    p.lastOrderDate || '',
    p.notes || ''
  ]);
  return response({ success: true, insertId: insertId });
}

function updateInsert(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(INSERTS_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(p.insertId)) {
      const row = i + 1;
      const updates = {
        2: p.description,    3: p.vendor,        4: p.productNumber,
        5: p.productUrl,     6: p.insertGeometry, 7: p.grade,
        8: p.materialApplication, 9: p.noseRadius, 10: p.icSize,
        11: p.qtyOnHand,    12: p.minQty,        13: p.lastPrice,
        14: p.lastOrderDate, 15: p.notes
      };
      Object.entries(updates).forEach(([col, val]) => {
        if (val !== undefined) sheet.getRange(row, Number(col)).setValue(val);
      });
      return response({ success: true });
    }
  }
  return response({ success: false, error: 'Insert not found: ' + p.insertId });
}

function updateInsertQty(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(INSERTS_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(p.insertId)) {
      sheet.getRange(i + 1, 11).setValue(Number(p.qty)); // col K = Qty On Hand
      return response({ success: true });
    }
  }
  return response({ success: false, error: 'Insert not found: ' + p.insertId });
}

// ─── Tooling: Holders ──────────────────────────────────────────────────────

function addHolder(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(HOLDERS_SHEET_NAME);
  // Duplicate check by product number
  if (p.productNumber && String(p.productNumber).trim() !== '') {
    const existing = sheet.getDataRange().getValues();
    for (var di = 1; di < existing.length; di++) {
      if (String(existing[di][3]).trim().toLowerCase() === String(p.productNumber).trim().toLowerCase()) {
        return response({ success: false, error: 'Duplicate: holder product # ' + p.productNumber + ' already exists as ' + existing[di][0] });
      }
    }
  }
  const holderId = generateId('HL', sheet);
  sheet.appendRow([
    holderId,
    p.description || '',
    p.vendor || '',
    p.productNumber || '',
    p.productUrl || '',
    p.type || '',
    p.taper || '',
    p.boreCapacity || '',
    p.qtyOnHand || 0,
    p.minQty || 0,
    p.lastPrice || '',
    p.lastOrderDate || '',
    p.notes || ''
  ]);
  return response({ success: true, holderId: holderId });
}

function updateHolder(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(HOLDERS_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(p.holderId)) {
      const row = i + 1;
      const updates = {
        2: p.description,  3: p.vendor,       4: p.productNumber,
        5: p.productUrl,   6: p.type,         7: p.taper,
        8: p.boreCapacity, 9: p.qtyOnHand,    10: p.minQty,
        11: p.lastPrice,   12: p.lastOrderDate, 13: p.notes
      };
      Object.entries(updates).forEach(([col, val]) => {
        if (val !== undefined) sheet.getRange(row, Number(col)).setValue(val);
      });
      return response({ success: true });
    }
  }
  return response({ success: false, error: 'Holder not found: ' + p.holderId });
}

// ─── Tooling: Accessories ──────────────────────────────────────────────────

function addAccessory(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(ACCESSORIES_SHEET_NAME);
  const accessoryId = generateId('AC', sheet);
  sheet.appendRow([
    accessoryId,
    p.toolId || '',
    p.type || '',
    p.description || '',
    p.vendor || '',
    p.productNumber || '',
    p.threadSize || '',
    p.driveType || '',
    p.notes || ''
  ]);
  return response({ success: true, accessoryId: accessoryId });
}

// ─── Tooling: Links ────────────────────────────────────────────────────────

function addInsertToolLink(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(INSERT_TOOL_LINKS_SHEET_NAME);
  // Check for duplicate
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]) === String(p.toolId) && String(values[i][2]) === String(p.insertId)) {
      return response({ success: false, error: 'Link already exists.' });
    }
  }
  const linkId = generateId('ITL', sheet);
  sheet.appendRow([linkId, p.toolId, p.insertId, p.notes || '']);
  return response({ success: true, linkId: linkId });
}

function removeInsertToolLink(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(INSERT_TOOL_LINKS_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]) === String(p.linkId)) {
      sheet.deleteRow(i + 1);
      return response({ success: true });
    }
  }
  return response({ success: false, error: 'Link not found.' });
}

function addToolPartLink(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TOOL_PART_LINKS_SHEET_NAME);
  // Check for duplicate
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]) === String(p.partNumber) && String(values[i][2]) === String(p.toolId)) {
      return response({ success: false, error: 'Link already exists.' });
    }
  }
  const linkId = generateId('TPL', sheet);
  sheet.appendRow([
    linkId,
    p.partNumber || '',
    p.toolId || '',
    p.fusionTNum || '',
    p.source || 'manual',
    p.fusionFile || '',
    p.notes || ''
  ]);
  return response({ success: true, linkId: linkId });
}

function removeToolPartLink(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TOOL_PART_LINKS_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]) === String(p.linkId)) {
      sheet.deleteRow(i + 1);
      return response({ success: true });
    }
  }
  return response({ success: false, error: 'Link not found.' });
}

// ─── Tooling: Machine Inventory ────────────────────────────────────────────

function updateToolMachineInventory(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TOOL_MACHINE_INV_SHEET_NAME);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]) === String(p.toolId) && String(values[i][2]) === String(p.machineId)) {
      // Update existing record
      sheet.getRange(i + 1, 4).setValue(Number(p.qty));
      sheet.getRange(i + 1, 5).setValue(p.inMagazine || 'N');
      return response({ success: true, recordId: values[i][0] });
    }
  }

  // Insert new record
  const recordId = generateId('TMI', sheet);
  sheet.appendRow([
    recordId,
    p.toolId,
    p.machineId,
    Number(p.qty) || 0,
    p.inMagazine || 'N'
  ]);
  return response({ success: true, recordId: recordId });
}

// ─── Tooling: Machines ─────────────────────────────────────────────────────

function addMachine(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(MACHINES_SHEET_NAME);
  const machineId = generateId('MCH', sheet);
  sheet.appendRow([machineId, p.name || '', p.type || '', p.notes || '']);
  return response({ success: true, machineId: machineId });
}

// ─── Reorder data ──────────────────────────────────────────────────────────

function getReorderData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const toolsSheet = ss.getSheetByName(TOOLS_SHEET_NAME);
  const insertsSheet = ss.getSheetByName(INSERTS_SHEET_NAME);
  const invSheet = ss.getSheetByName(TOOL_MACHINE_INV_SHEET_NAME);

  if (!toolsSheet || !insertsSheet || !invSheet) {
    return response({ success: false, error: 'Tooling sheets not initialized. Run initializeToolingSheets() first.' });
  }

  // Build machine qty map: toolId → total machine qty
  const invRows = invSheet.getDataRange().getValues().slice(1);
  const machineQtyMap = {};
  invRows.forEach(r => {
    const toolId = String(r[1]);
    const qty = Number(r[3]) || 0;
    machineQtyMap[toolId] = (machineQtyMap[toolId] || 0) + qty;
  });

  // Tools needing reorder: stockQty + machineQtys < minQty
  const toolRows = toolsSheet.getDataRange().getValues();
  const toolHeaders = toolRows[0];
  const reorderTools = toolRows.slice(1)
    .filter(r => r[0]) // skip empty rows
    .map(r => {
      const toolId   = String(r[0]);
      const stockQty = Number(r[17]) || 0; // col R
      const minQty   = Number(r[18]) || 0; // col S
      const machineQty = machineQtyMap[toolId] || 0;
      const totalQty = stockQty + machineQty;
      return {
        id:            toolId,
        category:      'Tool',
        description:   r[3],
        vendor:        r[4],
        productNumber: r[5],
        productUrl:    r[6],
        onHand:        totalQty,
        minQty:        minQty,
        lastPrice:     r[19],
        lastOrderDate: formatDate(r[20])
      };
    })
    .filter(t => t.onHand < t.minQty && t.minQty > 0);

  // Inserts needing reorder: qtyOnHand < minQty
  const insertRows = insertsSheet.getDataRange().getValues();
  const reorderInserts = insertRows.slice(1)
    .filter(r => r[0])
    .map(r => ({
      id:            String(r[0]),
      category:      'Insert',
      description:   r[1],
      vendor:        r[2],
      productNumber: r[3],
      productUrl:    r[4],
      onHand:        Number(r[10]) || 0,
      minQty:        Number(r[11]) || 0,
      lastPrice:     r[12],
      lastOrderDate: formatDate(r[13])
    }))
    .filter(i => i.onHand < i.minQty && i.minQty > 0);

  return response({
    success: true,
    reorderTools,
    reorderInserts
  });
}
