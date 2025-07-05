import { google } from 'googleapis';
import fs from 'fs';
import path, { dirname } from 'path';
import { fileURLToPath } from 'url';


// NEW (correct if stored in a folder inside your project)
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const keyFilePath = path.join(__dirname, 'credentials.json');

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

  let itemRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!A:A`,
  });

  let itemRows = itemRes.data.values || [];

  // Fill item list from template if sheet is empty
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

  const itemRowMap = {};
  itemRows.forEach((row, index) => {
    const name = row[0]?.trim().toLowerCase();
    if (name) itemRowMap[name] = index + 1;
  });

  const submittedItems = Object.entries(itemsMap).map(([name, qty]) => ({
    name,
    qty,
    normalized: name.trim().toLowerCase(),
  }));

  const matchedItems = submittedItems.filter(i => itemRowMap[i.normalized] !== undefined);
  if (matchedItems.length === 0) {
    throw new Error('None of the submitted items matched the sheet items.');
  }

  const colIndex = await getNextEmptyColumn(sheets, spreadsheetId, sheetName);
  const colLetter = getColumnLetter(colIndex);

  const maxRow = Math.max(...matchedItems.map(i => itemRowMap[i.normalized]));
  const values = Array(maxRow).fill(['']);
  values[0] = [date];
  matchedItems.forEach(i => {
    const row = itemRowMap[i.normalized];
    values[row - 1] = [i.qty || 0];
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!${colLetter}1:${colLetter}${maxRow}`,
    valueInputOption: 'RAW',
    requestBody: { values },
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

  const cols = res.data.values?.[0]?.length || 0;
  return cols + 1;
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
