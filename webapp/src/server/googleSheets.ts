import { google } from 'googleapis';

export interface GoogleSheetsRow {
  顧客番号: string;
  顧客名: string;
  車種名: string;
  シート名: string;
  ファイルパス: string;
}

export async function getGoogleSheetsData(): Promise<GoogleSheetsRow[]> {
  try {
    const apiKey = process.env.GOOGLE_SHEETS_API_KEY;
    const spreadsheetId = process.env.GOOGLE_SHEETS_SPREADSHEET_ID;
    
    if (!apiKey || !spreadsheetId) {
      throw new Error('Missing Google Sheets API configuration');
    }
    
    const sheets = google.sheets({ 
      version: 'v4', 
      auth: apiKey 
    });
    
    console.log('Fetching data from Google Sheets...');
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: spreadsheetId,
      range: 'A:E', // A列からE列まで
    });

    const rows = response.data.values || [];
    console.log(`Fetched ${rows.length} rows from Google Sheets`);

    // ヘッダー行をスキップ（1行目）
    const dataRows = rows.slice(1);
    
    const customers: GoogleSheetsRow[] = dataRows.map((row: string[]) => ({
      顧客番号: row[0] || '',
      顧客名: row[1] || '',
      車種名: row[2] || '',
      シート名: row[3] || '',
      ファイルパス: row[4] || '',
    }));

    console.log(`Converted ${customers.length} customers from Google Sheets`);
    return customers;
  } catch (error) {
    console.error('Error fetching Google Sheets data:', error);
    throw error;
  }
}
