import express from 'express';
import cors from 'cors';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

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

function readCustomers(): CustomerRow[] {
  try {
    // JSONファイルからデータを読み込み
    // Vercel環境では、dist/server/data/customers.json にある
    const dataPath = process.env.NODE_ENV === 'production' 
      ? path.join(__dirname, 'data', 'customers.json')
      : path.join(__dirname, 'data', 'customers.json');
    
    if (!fs.existsSync(dataPath)) {
      console.log('Customer data JSON file not found:', dataPath);
      console.log('Current directory:', __dirname);
      console.log('Available files:', fs.readdirSync(__dirname));
      return [];
    }
    
    console.log('Reading customer data from JSON file:', dataPath);
    const jsonData = fs.readFileSync(dataPath, 'utf8');
    const customers = JSON.parse(jsonData) as CustomerRow[];
    
    console.log(`Loaded ${customers.length} customers from JSON file`);
    return customers;
  } catch (error) {
    console.error('Error reading customer data:', error);
    return [];
  }
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

app.get('/api/customers', (_req, res) => {
  try {
    const rows = readCustomers();
    res.json({ count: rows.length, rows });
  } catch (err) {
    res.status(500).json({ error: 'Failed to read customer data', detail: String(err) });
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
