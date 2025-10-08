import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';
import { getCustomerSheetValues } from './googleSheets.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors());
app.use(express.json());

const staticDirCandidates = [
  path.resolve(process.cwd(), 'dist', 'public'),
  path.resolve(process.cwd(), 'src', 'public'),
  path.resolve(__dirname, '..', 'public'),
  path.resolve(__dirname, 'public')
];

for (const dir of staticDirCandidates) {
  if (fs.existsSync(dir)) {
    app.use(express.static(dir));
  }
}

app.get('/health', (_req, res) => {
  res.json({ ok: true });
});

type CustomerRow = {
  顧客番号: string;
  顧客名: string;
  車種名: string;
  シート名: string;
  ファイルパス: string;
  フォルダ名: string;
  顧客コード: string;
  '顧客名（フォルダ）': string;
  '名　前': string;
  'ふりがな': string;
  '自宅郵便番号': string;
  '自宅住所1': string;
  '自宅住所2': string;
  '携帯番号': string;
  '自宅電話番号': string;
  '会社名': string;
  '会社郵便番号': string;
  '車　名': string;
  '車両ナンバー': string;
  '車台番号': string;
  '型式・年式番号': string;
  '類別番号': string;
  '初年度': string;
  '次回車検満期日': string;
};

const FIELD_MAP: Record<string, keyof CustomerRow> = {
  'コード№': '顧客番号',
  '名　前': '名　前',
  'ふりがな': 'ふりがな',
  '自宅郵便番号': '自宅郵便番号',
  '自宅住所1': '自宅住所1',
  '自宅住所2': '自宅住所2',
  '携帯番号': '携帯番号',
  '自宅電話番号': '自宅電話番号',
  '会社名': '会社名',
  '会社郵便番号': '会社郵便番号',
  '車　名': '車　名',
  '車両ナンバー': '車両ナンバー',
  '車台番号': '車台番号',
  '型式・年式番号': '型式・年式番号',
  '類別番号': '類別番号',
  '初年度': '初年度',
  '次回車検満期日': '次回車検満期日',
};

function handleField_(customer: CustomerRow, label: string, value: unknown) {
  if (!label) return;
  const key = FIELD_MAP[label];
  if (!key) return;

  let stringValue = '';
  if (value == null || value === '') {
    stringValue = '';
  } else {
    if (key === '次回車検満期日') {
      stringValue = normalizeExcelDate(value);
    } else if (key === '初年度') {
      stringValue = normalizeExcelDate(value, true);
    } else {
      stringValue = String(value);
    }
  }

  if (key === '次回車検満期日') {
    customer[key] = stringValue;
    return;
  }

  customer[key] = stringValue;
}

// 統合Excelブックからデータを取得する関数
async function readCustomers(): Promise<CustomerRow[]> {
  try {
    const sheetValues = await getCustomerSheetValues();

    console.log(`Google Sheetsから ${sheetValues.length} 件の顧客シートを処理中...`);

    const customers: CustomerRow[] = [];
    const processedCustomers = new Set<string>();

    for (const sheet of sheetValues) {
      const data = sheet.values;
      if (data.length < 2) continue;

      const folderInfo = parseFolderName(sheet.sheetName);
      const customerKey = folderInfo.code || sheet.sheetName;
      if (processedCustomers.has(customerKey)) continue;
      processedCustomers.add(customerKey);

      const customer: CustomerRow = {
        顧客番号: folderInfo.code,
        顧客名: folderInfo.name,
        車種名: '',
        シート名: sheet.sheetName,
        ファイルパス: `GoogleSheet: ${sheet.sheetName}`,
        フォルダ名: folderInfo.name,
        顧客コード: folderInfo.code,
        '顧客名（フォルダ）': folderInfo.name,
        '名　前': folderInfo.name,
        'ふりがな': '',
        '自宅郵便番号': '',
        '自宅住所1': '',
        '自宅住所2': '',
        '携帯番号': '',
        '自宅電話番号': '',
        '会社名': '',
        '会社郵便番号': '',
        '車　名': '',
        '車両ナンバー': '',
        '車台番号': '',
        '型式・年式番号': '',
        '類別番号': '',
        '初年度': '',
        '次回車検満期日': '',
      };

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row) continue;
        const leftLabel = String(row[0] ?? '').trim();
        const leftValue = row[1];
        handleField_(customer, leftLabel, leftValue);

        const rightLabel = String(row[2] ?? '').trim();
        const rightValue = row[3];
        handleField_(customer, rightLabel, rightValue);
      }

      // 名前・ふりがなの補完
      if (!customer['名　前']) customer['名　前'] = folderInfo.name;
      if (!customer['顧客名']) customer['顧客名'] = customer['名　前'];
      if (!customer['車　名']) customer['車　名'] = folderInfo.car || '';
      if (!customer['ふりがな']) {
        const parsed = parseCustomerName(customer['名　前']);
        customer['ふりがな'] = parsed.kana;
      }

      customers.push(customer);
    }
    
    console.log(`Google Sheetsから ${customers.length} 件の顧客データを取得しました`);
    return customers;
  } catch (error) {
    console.error('Google Sheetsからの顧客データ読み込みエラー:', error);
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

function normalizeExcelDate(value: unknown, isFirstRegistration = false): string {
  if (value == null || value === '') return '';

  if (typeof value === 'number') {
    return convertExcelDateToString(value);
  }

  if (value instanceof Date) {
    return formatDate(value);
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return '';

    // Already formatted YYYY/MM/DD or YYYY-MM-DD
    if (/^\d{4}[\.\/\-]\d{1,2}[\.\/\-]\d{1,2}$/.test(trimmed)) {
      const d = new Date(trimmed.replace(/[-.]/g, '/'));
      if (!Number.isNaN(d.getTime())) return formatDate(d);
    }

    // Pure numeric string
    if (/^\d+$/.test(trimmed)) {
      const num = Number(trimmed);
      if (!Number.isNaN(num)) {
        // Treat up to 5 digits (or < 60000) as Excel serial
        if (trimmed.length <= 5 || num < 60000) {
          return convertExcelDateToString(num);
        }
        // Treat 8 digit numbers as YYYYMMDD
        if (trimmed.length === 8) {
          const year = Number(trimmed.slice(0, 4));
          const month = Number(trimmed.slice(4, 6));
          const day = Number(trimmed.slice(6, 8));
          if (!Number.isNaN(year) && !Number.isNaN(month) && !Number.isNaN(day)) {
            const d = new Date(year, month - 1, day);
            if (!Number.isNaN(d.getTime())) return formatDate(d);
          }
        }
      }
    }

    // Japanese era strings etc. are returned as-is for 初年度
    if (isFirstRegistration) {
      return trimmed;
    }

    return trimmed;
  }

  return '';
}

function formatDate(date: Date): string {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  return `${year}/${month}/${day}`;
}

// フォルダ名から顧客情報を抽出する関数
function parseFolderName(folderName: string): { code: string; name: string; car?: string | undefined } {
  // 例: "0235となか戸中公子様(ﾐﾗｲｰｽ)"
  const match = folderName.match(/^(\d+)(.+)$/);
  if (match && match[1] && match[2]) {
    const rest = match[2];
    const carMatch = rest.match(/（([^）]+)）|\(([^)]+)\)$/);
    const carName = carMatch ? carMatch[1] || carMatch[2] : undefined;
    const name = rest.replace(/[()（）]/g, '').replace(/様.*$/, '').trim();
    return {
      code: match[1],
      name: name || rest.trim(),
      car: carName ?? undefined,
    };
  }
  return { code: '', name: folderName };
}

// 顧客名から名前とふりがなを分離する関数
function parseCustomerName(customerName: string): { name: string; kana: string } {
  if (!customerName) return { name: '', kana: '' };
  
  // 例: "おおひら大平恵美" → "大平恵美", "おおひら"
  // ひらがな部分と漢字部分を分離
  const kanaMatch = customerName.match(/^([あ-ん]+)/);
  const kanjiMatch = customerName.match(/([一-龯]+.*)$/);
  
      if (kanaMatch && kanjiMatch) {
        return {
          kana: kanaMatch[1] || '',
          name: kanjiMatch[1] || ''
        };
      }
  
  // 分離できない場合は、そのまま名前として使用
  return {
    name: customerName,
    kana: ''
  };
}


app.get('/api/customers', async (_req, res) => {
  try {
    const rows = await readCustomers();
    res.json({ count: rows.length, rows });
  } catch (err) {
    console.error('API Error:', err);
    res.status(500).json({ error: '統合Excelブックからの顧客データ読み込みに失敗しました', detail: String(err) });
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
