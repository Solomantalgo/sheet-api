import { google } from 'googleapis';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// ✅ Use secret file path from Render
const keyFilePath = '/etc/secrets/GOOGLE_CREDENTIALS_FILE';

// ✅ Check if credentials file exists
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
};

const DEFAULT_TEMPLATE_TAB = 'Acacia';

export async function appendReport(merchandiser, outlet, date, itemsMap) {
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
    if (name) itemRowMap[name] = index + 1; // sheet rows are 1-based
  });

  // Parse itemsMap into array with all needed fields and fix qty parsing
  const submittedItems = Object.entries(itemsMap).map(([name, item]) => {
    let qty = item.qty;
    if (qty === null || qty === undefined || qty === 'null' || qty === '') qty = 0;
    else if (typeof qty !== 'number') qty = parseInt(qty, 10);
    if (isNaN(qty)) qty = 0;

    let expiry = item.expiry;
    if (expiry === null || expiry === undefined || expiry === 'null') expiry = '';

    let notes = item.notes;
    if (notes === null || notes === undefined || notes === 'null') notes = '';

    console.log(`DEBUG: Parsed item -> Name: ${name}, Qty: ${qty}, Expiry: ${expiry}, Notes: ${notes}`);

    return {
      name,
      qty,
      expiry,
      notes,
      normalized: name.trim().toLowerCase(),
    };
  });

  const matchedItems = submittedItems.filter(i => itemRowMap[i.normalized] !== undefined);
  if (matchedItems.length === 0) {
    throw new Error('None of the submitted items matched the sheet items.');
  }

  const colIndex = await getNextEmptyColumn(sheets, spreadsheetId, sheetName);
  const qtyCol = getColumnLetter(colIndex);
  const expiryCol = getColumnLetter(colIndex + 1);
  const notesCol = getColumnLetter(colIndex + 2);

  const totalRows = itemRows.length;
  const startRow = 6; // Start writing qty/expiry/notes from row 6
  const numRowsToWrite = totalRows - startRow + 1;

  // Prepare arrays filled with empty strings for qty, expiry, notes
  const qtyValues = Array(numRowsToWrite).fill(['']);
  const expiryValues = Array(numRowsToWrite).fill(['']);
  const notesValues = Array(numRowsToWrite).fill(['']);

  // Fill arrays according to item rows
  matchedItems.forEach(i => {
    const row = itemRowMap[i.normalized];
    if (row >= startRow) {
      const arrIndex = row - startRow; // zero-based index for arrays
      qtyValues[arrIndex] = [i.qty];
      expiryValues[arrIndex] = [i.expiry];
      notesValues[arrIndex] = [i.notes];
      console.log(`DEBUG: Writing to row ${row} (array index ${arrIndex}) - Qty: ${i.qty}, Expiry: ${i.expiry}, Notes: ${i.notes}`);
    }
  });

  // 1. Write Date at row 1 of quantity column
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!${qtyCol}1`,
    valueInputOption: 'RAW',
    requestBody: { values: [[date]] },
  });

  // 2. Write Quantity values starting from row 6
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!${qtyCol}${startRow}:${qtyCol}${totalRows}`,
    valueInputOption: 'RAW',
    requestBody: { values: qtyValues },
  });

  // 3. Write "Expiry" header at row 1 of expiry column
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!${expiryCol}1`,
    valueInputOption: 'RAW',
    requestBody: { values: [['Expiry']] },
  });

  // 4. Write Expiry values starting from row 6
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!${expiryCol}${startRow}:${expiryCol}${totalRows}`,
    valueInputOption: 'RAW',
    requestBody: { values: expiryValues },
  });

  // 5. Write "Notes" header at row 1 of notes column
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!${notesCol}1`,
    valueInputOption: 'RAW',
    requestBody: { values: [['Notes']] },
  });

  // 6. Write Notes values starting from row 6
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!${notesCol}${startRow}:${notesCol}${totalRows}`,
    valueInputOption: 'RAW',
    requestBody: { values: notesValues },
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
  // Find the last non-empty cell in row 1
  let lastCol = 0;
  for (let i = 0; i < row.length; i++) {
    if (row[i] && row[i].toString().trim() !== '') lastCol = i + 1;
  }
  return lastCol + 1; // Next empty column (1-based)
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