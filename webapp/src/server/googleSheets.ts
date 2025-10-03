import 'dotenv/config';
import { google } from 'googleapis';

export interface GoogleSheetValues {
  sheetName: string;
  values: (string | number | undefined)[][];
}

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

  const range = "A1:D50";

  const results: GoogleSheetValues[] = [];

  for (const name of sheetNames) {
    try {
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `'${name}'!${range}`,
        majorDimension: 'ROWS',
      });
      const values = (res.data.values as (string | number | undefined)[][]) || [];
      results.push({ sheetName: name, values });
    } catch (error) {
      console.error(`Error fetching sheet ${name}:`, error);
    }
  }

  return results;
}
