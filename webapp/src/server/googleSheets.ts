import 'dotenv/config';
import { google } from 'googleapis';

export interface GoogleSheetValues {
  sheetName: string;
  values: (string | number | undefined)[][];
}

const BATCH_SIZE = 40; // 40 ranges ≒ 40 sheets / request (well under quota limits)
const RANGE = 'A1:D50';

export async function getCustomerSheetValues(): Promise<GoogleSheetValues[]> {
  const apiKey = process.env.GOOGLE_SHEETS_API_KEY;
  const spreadsheetId = process.env.GOOGLE_SHEETS_SPREADSHEET_ID;

  if (!apiKey || !spreadsheetId) {
    throw new Error('Missing Google Sheets API configuration');
  }

  const sheets = google.sheets({ version: 'v4', auth: apiKey });

  console.log('Fetching sheet metadata from Google Sheets...');

  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'sheets(properties(title))',
  });

  const sheetNames = (meta.data.sheets || [])
    .map((sheet) => sheet.properties?.title || '')
    .filter((title) => title.includes('_顧客名簿'));

  console.log(`Found ${sheetNames.length} 顧客名簿 sheets`);

  const results: GoogleSheetValues[] = [];

  for (let i = 0; i < sheetNames.length; i += BATCH_SIZE) {
    const chunkNames = sheetNames.slice(i, i + BATCH_SIZE);
    const ranges = chunkNames.map((name) => `'${escapeSheetName(name)}'!${RANGE}`);

    try {
      const response = await sheets.spreadsheets.values.batchGet({
        spreadsheetId,
        ranges,
        majorDimension: 'ROWS',
      });

      const valueRanges = response.data.valueRanges || [];

      chunkNames.forEach((name, idx) => {
        const values = (valueRanges[idx]?.values as (string | number | undefined)[][]) || [];
        results.push({ sheetName: name, values });
      });
    } catch (error) {
      console.error('Error fetching sheet chunk:', error);
    }
  }

  return results;
}

function escapeSheetName(name: string): string {
  return name.replace(/'/g, "''");
}
