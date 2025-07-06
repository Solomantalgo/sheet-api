import { google } from 'googleapis';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const keyFilePath = '/etc/secrets/GOOGLE_CREDENTIALS_FILE';

if (!fs.existsSync(keyFilePath)) {
  throw new Error(`❌ Google credentials file not found at: ${keyFilePath}`);
}

const auth = new google.auth.GoogleAuth({
  keyFile: keyFilePath,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const SPREADSHEET_IDS = {
  Solomon: '148PXW2iApr04lOo-3rY4a8IJ2ToUztMDOzN6TceC3bQ',
  Patricia: '1sm014ASh1w84UJAfXj7OX7Gtj29Mv5e1xin_eJ5Pz78',
  Milan: '1HwoB4FFQlqkgwJusbZxbBUaB861p9flE6hMtNk6e-Tw',
  Caro: '1GDc1MLSVRwm_Ccy4ahkVeNpfTzlLH6iOIRZYQ6PfbJU',
  Charles: '1Ji2e5ewH8bVieMKKR57Xd2ZBYUwT5XfmA6Qk3uH_qDo',
  Brenda: '1EPdIDDTgmauEbTlSeme9bxKNUjN9i1RTC-RZLCKNLOQ',
  Rayan: '1uwE2JTyPxoxrvHvebhWfwmvZNy3mlrLyvEatn6JyHm8',
  Job: '1yUySiotqekdK5EqcO_A637qwDg0PdU0P0Kb8_9WTAns',
  Test: '1TEVqW_5TjGFp-ibqg97dgec1qTULoBkGPl35_3i4nz4',
};

const DEFAULT_TEMPLATE_TAB = 'Acacia';

export async function appendReport(merchandiser, outlet, date, notes, items) {
  const spreadsheetId = SPREADSHEET_IDS[merchandiser];
  if (!spreadsheetId) throw new Error(`Spreadsheet not found for merchandiser: ${merchandiser}`);

  const authClient = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  const sheetName = await ensureSheetExists(sheets, spreadsheetId, outlet, DEFAULT_TEMPLATE_TAB);

  // Get all item names from column A
  let itemRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!A:A`,
  });
  let itemRows = itemRes.data.values || [];

  // If empty, fill from template
  if (itemRows.length === 0 || itemRows.every(row => !row[0])) {
    const templateRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${DEFAULT_TEMPLATE_TAB}!A:A`,
    });
    const templateItems = templateRes.data.values || [];

    if (templateItems.length > 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!A1:A${templateItems.length}`,
        valueInputOption: 'RAW',
        requestBody: { values: templateItems },
      });
      itemRows = templateItems;
    }
  }

  // Map item names to their row number (1-based)
  const itemRowMap = {};
  itemRows.forEach((row, index) => {
    const name = row[0]?.trim().toLowerCase();
    if (name) itemRowMap[name] = index + 1;
  });

  // Debug log for incoming items
  console.log('DEBUG: items received:', JSON.stringify(items, null, 2));

  // Support both array and object formats for items (for backward compatibility)
  let submittedItems = [];
  if (Array.isArray(items)) {
    submittedItems = items.map(item => {
      let qty = item.qty;
      if (qty === null || qty === undefined || qty === 'null' || qty === '') qty = 0;
      else if (typeof qty !== 'number') qty = parseInt(qty, 10);
      if (isNaN(qty)) qty = 0;

      let expiry = item.expiry;
      if (expiry === null || expiry === undefined || expiry === 'null' || expiry === '') expiry = 'Null';

      const name = item.name;
      console.log(`DEBUG: Parsed item -> Name: ${name}, Qty: ${qty}, Expiry: ${expiry}`);

      return {
        name,
        qty,
        expiry,
        normalized: name.trim().toLowerCase(),
      };
    });
  } else {
    submittedItems = Object.entries(items).map(([name, item]) => {
      let qty = item.qty;
      if (qty === null || qty === undefined || qty === 'null' || qty === '') qty = 0;
      else if (typeof qty !== 'number') qty = parseInt(qty, 10);
      if (isNaN(qty)) qty = 0;

      let expiry = item.expiry;
      if (expiry === null || expiry === undefined || expiry === 'null' || expiry === '') expiry = 'Null';

      console.log(`DEBUG: Parsed item -> Name: ${name}, Qty: ${qty}, Expiry: ${expiry}`);

      return {
        name,
        qty,
        expiry,
        normalized: name.trim().toLowerCase(),
      };
    });
  }

  const matchedItems = submittedItems.filter(i => itemRowMap[i.normalized] !== undefined);
  if (matchedItems.length === 0) {
    throw new Error('None of the submitted items matched the sheet items.');
  }

  const colIndex = await getNextEmptyColumn(sheets, spreadsheetId, sheetName);
  const qtyCol = getColumnLetter(colIndex);
  const expiryCol = getColumnLetter(colIndex + 1);

  // Ensure enough columns exist for qty and expiry only
  await ensureEnoughColumns(sheets, spreadsheetId, sheetName, colIndex + 1);

  const totalRows = itemRows.length;
  const startRow = 6;
  const numRowsToWrite = totalRows - startRow + 1;

  // Prepare arrays filled with empty strings for qty and expiry
  const qtyValues = Array(numRowsToWrite).fill(['']);
  const expiryValues = Array(numRowsToWrite).fill(['']);

  matchedItems.forEach(i => {
    const row = itemRowMap[i.normalized];
    if (row >= startRow) {
      const arrIndex = row - startRow;
      qtyValues[arrIndex] = [i.qty];
      expiryValues[arrIndex] = [i.expiry];
      console.log(`DEBUG: Writing to row ${row} (array index ${arrIndex}) - Qty: ${i.qty}, Expiry: ${i.expiry}`);
    }
  });

  // 1. Write Date at row 1 of quantity column
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!${qtyCol}1`,
    valueInputOption: 'RAW',
    requestBody: { values: [[date]] },
  });

  // 2. Write notes just below the date column in row 2 (from frontend, or fallback)
  const notesValue = notes == null || notes === '' ? 'No notes for this day' : notes;
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!${qtyCol}2`,
    valueInputOption: 'RAW',
    requestBody: { values: [[notesValue]] },
  });

  // 3. Write Quantity values starting from row 6
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!${qtyCol}${startRow}:${qtyCol}${totalRows}`,
    valueInputOption: 'RAW',
    requestBody: { values: qtyValues },
  });

  // 4. Write "Expiry" header at row 1 of expiry column
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!${expiryCol}1`,
    valueInputOption: 'RAW',
    requestBody: { values: [['Expiry']] },
  });

  // 5. Write Expiry values starting from row 6
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!${expiryCol}${startRow}:${expiryCol}${totalRows}`,
    valueInputOption: 'RAW',
    requestBody: { values: expiryValues },
  });

  // --- FORMATTING SECTION ---

  function columnLetterToIndex(letter) {
    let col = 0;
    for (let i = 0; i < letter.length; i++) {
      col *= 26;
      col += letter.charCodeAt(i) - 65 + 1;
    }
    return col - 1;
  }
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = spreadsheet.data.sheets.find(s => s.properties.title === sheetName);
  const sheetId = sheet.properties.sheetId;
  const dateColIndex = columnLetterToIndex(qtyCol);
  const expiryColIndex = columnLetterToIndex(expiryCol);

// --- SINGLE NOTE ON NOTES CELL (row 2, date column) ---
const noteRequest = {
  updateCells: {
    rows: [
      {
        values: [
          {
            note: notesValue,
            userEnteredFormat: {
              textFormat: {
                fontSize: 10, // readable font size
              },
            },
          },
        ],
      },
    ],
    fields: 'note,userEnteredFormat.textFormat.fontSize',
    start: {
      sheetId: sheetId,
      rowIndex: 1, // Row 2 (0-indexed)
      columnIndex: dateColIndex, // Matches where the notes are written
    },
  },
};


  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        // Limit row height for notes (row 2)
        noteRequest, // 👈 Insert the note formatting FIRST
{
  updateDimensionProperties: {
    range: {
      sheetId: sheetId,
      dimension: 'ROWS',
      startIndex: 1, // Row 2
      endIndex: 2,
    },
    properties: {
      pixelSize: 60, // You can increase or decrease this to control visible height
    },
    fields: 'pixelSize',
  },
},
        // Notes cell (row 2, date column): text wrap + light yellow fill + dark blue text
        {
          repeatCell: {
            range: {
              sheetId: sheetId,
              startRowIndex: 1, // Row 2 (zero-based)
              endRowIndex: 2,
              startColumnIndex: dateColIndex,
              endColumnIndex: dateColIndex + 1,
            },
            cell: {
              userEnteredFormat: {
                wrapStrategy: 'WRAP',
                backgroundColor: {
                  red: 1,
                  green: 1,
                  blue: 0.6, // Light yellow
                },
                textFormat: {
                  foregroundColor: { red: 0.1, green: 0.2, blue: 0.5 }, // dark blue text
                },
              },
            },
            fields: 'userEnteredFormat(wrapStrategy,backgroundColor,textFormat.foregroundColor)',
          },
        },
        // Date header (row 1): blue text, light blue background, bold
        {
          repeatCell: {
            range: {
              sheetId: sheetId,
              startRowIndex: 0, // Row 1 (zero-based)
              endRowIndex: 1,
              startColumnIndex: dateColIndex,
              endColumnIndex: dateColIndex + 1,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.8, green: 0.9, blue: 1 },
                textFormat: { foregroundColor: { red: 0.1, green: 0.2, blue: 0.8 }, bold: true },
                wrapStrategy: 'WRAP',
              },
            },
            fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy)',
          },
        },
        // Date values (row 6+): light blue background, dark blue text, bold
        {
          repeatCell: {
            range: {
              sheetId: sheetId,
              startRowIndex: 2, // Skip notes row
              endRowIndex: totalRows,
              startColumnIndex: dateColIndex,
              endColumnIndex: dateColIndex + 1,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.8, green: 0.9, blue: 1 },
                textFormat: { foregroundColor: { red: 0.1, green: 0.2, blue: 0.5 }, bold: true },
                wrapStrategy: 'WRAP',
              },
            },
            fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy)',
          },
        },
        // Expiry header (row 1): red text, light green background, bold
        {
          repeatCell: {
            range: {
              sheetId: sheetId,
              startRowIndex: 0, // Row 1 (zero-based)
              endRowIndex: 1,
              startColumnIndex: expiryColIndex,
              endColumnIndex: expiryColIndex + 1,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.85, green: 1, blue: 0.85 },
                textFormat: { foregroundColor: { red: 0.8, green: 0.1, blue: 0.1 }, bold: true },
                wrapStrategy: 'WRAP',
              },
            },
            fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy)',
          },
        },
        // Expiry values (row 2+): light green background, dark green text, bold
        {
          repeatCell: {
            range: {
              sheetId: sheetId,
              startRowIndex: 1, // From row 2 (expiry values)
              endRowIndex: totalRows,
              startColumnIndex: expiryColIndex,
              endColumnIndex: expiryColIndex + 1,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.85, green: 1, blue: 0.85 },
                textFormat: { foregroundColor: { red: 0.1, green: 0.4, blue: 0.1 }, bold: true },
                wrapStrategy: 'WRAP',
              },
            },
            fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy)',
          },
        },
      ],
    },
  });

  console.log(`✅ Report saved: ${merchandiser} > ${sheetName}`);
}

async function ensureSheetExists(sheets, spreadsheetId, outlet, templateTab) {
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
  const existingTabs = spreadsheet.data.sheets.map(s => s.properties.title);
  const cleanName = outlet.trim();

  if (existingTabs.includes(cleanName)) {
    return cleanName;
  }

  const templateSheet = spreadsheet.data.sheets.find(s => s.properties.title === templateTab);
  if (!templateSheet) throw new Error(`Template tab "${templateTab}" not found`);

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          duplicateSheet: {
            sourceSheetId: templateSheet.properties.sheetId,
            newSheetName: cleanName,
          },
        },
      ],
    },
  });

  console.log(`🆕 Created new tab "${cleanName}" from template "${templateTab}"`);
  return cleanName;
}

async function getNextEmptyColumn(sheets, spreadsheetId, sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!1:1`,
  });

  const row = res.data.values?.[0] || [];
  let lastCol = 0;
  for (let i = 0; i < row.length; i++) {
    if (row[i] && row[i].toString().trim() !== '') lastCol = i + 1;
  }
  return lastCol + 1;
}

function getColumnLetter(colNum) {
  let letter = '';
  while (colNum > 0) {
    let mod = (colNum - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    colNum = Math.floor((colNum - 1) / 26);
  }
  return letter;
}

// Helper to ensure enough columns exist
async function ensureEnoughColumns(sheets, spreadsheetId, sheetName, neededColIndex) {
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = spreadsheet.data.sheets.find(s => s.properties.title === sheetName);
  const currentCols = sheet.properties.gridProperties.columnCount;

  if (currentCols < neededColIndex) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            appendDimension: {
              sheetId: sheet.properties.sheetId,
              dimension: 'COLUMNS',
              length: neededColIndex - currentCols,
            },
          },
        ],
      },
    });
    console.log(`Added ${neededColIndex - currentCols} columns to "${sheetName}"`);
  }
}

//