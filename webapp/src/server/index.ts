import express from 'express';
import cors from 'cors';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';
import { getGoogleSheetsData } from './googleSheets.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors());
app.use(express.json());

// Serve static frontend
if (process.env.NODE_ENV === 'production') {
  // Vercel環境では、プロジェクトルートからの相対パスを使用
  app.use(express.static(path.resolve(process.cwd(), 'webapp', 'src', 'public')));
} else {
  // 開発環境
  app.use(express.static(path.resolve(__dirname, '..', 'public')));
  app.use(express.static(path.resolve(__dirname, 'public')));
}

app.get('/health', (_req, res) => {
  res.json({ ok: true });
});

// Load Excel from configured path
function getExcelPath(): string {
  // default to sibling folder "0顧客マスター/顧客ファイル.xlsx" at workspace root
  // Allow override via env EXCEL_PATH
  const envPath = process.env.EXCEL_PATH;
  if (envPath && fs.existsSync(envPath)) return envPath;

  // try absolute known path based on provided workspace
  const defaultPath = path.resolve(
    __dirname,
    '..', // server -> parent
    '..', // -> webapp (project root)
    '..', // -> workspace root
    '0顧客マスター',
    '顧客ファイル.xlsx'
  );
  return defaultPath;
}

type CustomerRow = {
  名前?: string;
  振り仮名?: string;
  郵便番号?: string;
  自宅住所?: string;
  自宅住所2?: string;
  携帯番号?: string;
  自宅電話番号?: string;
  会社名?: string;
  会社の郵便番号?: string;
  社名?: string;
  車両ナンバー?: string;
  車台番号?: string;
  型式?: string;
  類別番号?: string;
  初年度?: string | number | Date;
  次回車検満期日?: string | number | Date;
  [key: string]: unknown;
};

// Google Sheets APIを使用してデータを取得する関数
async function readCustomers(): Promise<CustomerRow[]> {
  try {
    const googleSheetsData = await getGoogleSheetsData();
    
    // Google SheetsデータをCustomerRow形式に変換
    const customers: CustomerRow[] = googleSheetsData.map(row => {
      // シート名から顧客情報を抽出
      const folderInfo = parseFolderName(row.シート名);
      
      // ファイルパスからフォルダ名を抽出
      const folderName = extractFolderNameFromPath(row.ファイルパス);
      
      return {
        '顧客番号': row.顧客番号,
        '顧客名': row.顧客名,
        '車種名': row.車種名,
        'シート名': row.シート名,
        'ファイルパス': row.ファイルパス,
        'フォルダ名': folderName,
        '顧客コード': folderInfo.code,
        '顧客名（フォルダ）': folderInfo.name,
        // 既存のCustomerRow形式に合わせて追加フィールドを設定
        '名　前': row.顧客名,
        'ふりがな': '',
        '自宅郵便番号': '',
        '自宅住所1': '',
        '自宅住所2': '',
        '次回車検満期日': '',
        '初年度': '',
      } as CustomerRow;
    });
    
    console.log(`Converted ${customers.length} customers from Google Sheets`);
    return customers;
  } catch (error) {
    console.error('Error reading customer data from Google Sheets:', error);
    return [];
  }
}

// ファイルパスからフォルダ名を抽出する関数
function extractFolderNameFromPath(filePath: string): string {
  const pathParts = filePath.split('/');
  const folderPart = pathParts.find(part => part.includes('様'));
  return folderPart || '';
}

// Excel日付シリアル値を文字列に変換する関数
function convertExcelDateToString(excelDate: number): string {
  // Excelの日付シリアル値（1900年1月1日からの日数）
  // Excelの日付システムのバグにより、1900年は閏年として扱われるため、-2を調整
  const excelEpoch = new Date(1900, 0, 1);
  const actualDate = new Date(excelEpoch.getTime() + (excelDate - 2) * 24 * 60 * 60 * 1000);
  
  // 日本語形式で日付をフォーマット
  const year = actualDate.getFullYear();
  const month = (actualDate.getMonth() + 1).toString().padStart(2, '0');
  const day = actualDate.getDate().toString().padStart(2, '0');
  
  return `${year}/${month}/${day}`;
}

// フォルダ名から顧客情報を抽出する関数
function parseFolderName(folderName: string): { code: string; name: string } {
  // 例: "0235となか戸中公子様(ﾐﾗｲｰｽ)"
  const match = folderName.match(/^(\d+)(.+)$/);
  if (match && match[1] && match[2]) {
    return {
      code: match[1],
      name: match[2].replace(/[()（）]/g, '') // 括弧を除去
    };
  }
  return { code: '', name: folderName };
}

app.get('/api/customers', async (_req, res) => {
  try {
    const rows = await readCustomers();
    res.json({ count: rows.length, rows });
  } catch (err) {
    res.status(500).json({ error: 'Failed to read customer data from Google Sheets', detail: String(err) });
  }
});

// Vercel環境では app.listen() は不要
if (process.env.NODE_ENV !== 'production') {
  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () => {
    // eslint-disable-next-line no-console
    console.log(`Server listening on http://localhost:${PORT}`);
  });
}

export default app;
