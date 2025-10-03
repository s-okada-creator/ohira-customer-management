import express from 'express';
import cors from 'cors';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';
import xlsx from 'xlsx';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors());
app.use(express.json());

// Serve static frontend (dev: src/public, prod: dist/public)
app.use(express.static(path.resolve(__dirname, '..', 'public')));
app.use(express.static(path.resolve(__dirname, 'public')));

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
  const customersFolder = path.resolve(
    __dirname,
    '..', // server -> src
    '..', // src -> webapp
    '..', // webapp -> workspace root
    '☆顧客フォルダ'
  );

  if (!fs.existsSync(customersFolder)) {
    console.log('Customers folder not found:', customersFolder);
    return [];
  }

  const allCustomers: CustomerRow[] = [];
  
  try {
    // 顧客フォルダ内のすべてのサブフォルダを取得
    const customerFolders = fs.readdirSync(customersFolder, { withFileTypes: true })
      .filter(dirent => dirent.isDirectory())
      .map(dirent => dirent.name);

    console.log(`Found ${customerFolders.length} customer folders`);

    for (const folderName of customerFolders) {
      const customerFilePath = path.join(customersFolder, folderName, '顧客ファイル.xlsx');
      
      if (fs.existsSync(customerFilePath)) {
        try {
          const workbook = xlsx.readFile(customerFilePath);
          
          // 顧客データベースシートを優先的に探す
          let targetSheet = null;
          for (const sheetName of workbook.SheetNames) {
            if (sheetName === '顧客データベース') {
              targetSheet = workbook.Sheets[sheetName];
              break;
            }
          }
          
          // 顧客データベースシートが見つからない場合は最初のシートを使用
          if (!targetSheet && workbook.SheetNames.length > 0) {
            targetSheet = workbook.Sheets[workbook.SheetNames[0]];
          }
          
          if (targetSheet) {
            // シートの範囲を確認
            const range = targetSheet['!ref'];
            if (range) {
              // 行と列の数を取得
              const rangeParts = range.split(':');
              const endCell = rangeParts[1];
              const endRow = parseInt(endCell.replace(/[A-Z]/g, ''));
              
              // 2行目以降のデータを取得（1行目はヘッダー）
              if (endRow >= 2) {
                // ヘッダー行を取得
                const headerRow = xlsx.utils.sheet_to_json(targetSheet, { header: 1, range: 0 })[0];
                
                // データ行を取得（2行目から）
                const dataRows = xlsx.utils.sheet_to_json(targetSheet, { header: 1, range: 1 });
                
                // 各有効なデータ行を処理
                dataRows.forEach(dataRow => {
                  // データ行が空でないかチェック
                  if (dataRow.length > 0 && dataRow[1] && String(dataRow[1]).trim() !== '') {
                    // ヘッダーとデータを組み合わせてオブジェクトを作成
                    const customerData: CustomerRow = {};
                    
                    headerRow.forEach((header, index) => {
                      if (header && dataRow[index] !== undefined) {
                        customerData[header] = dataRow[index];
                      }
                    });
                    
                    // フォルダ名から顧客情報を抽出
                    const folderInfo = parseFolderName(folderName);
                    
                    // 日付フィールドを変換
                    if (customerData['次回車検満期日'] && typeof customerData['次回車検満期日'] === 'number') {
                      customerData['次回車検満期日'] = convertExcelDateToString(customerData['次回車検満期日']);
                    }
                    if (customerData['初年度'] && typeof customerData['初年度'] === 'number') {
                      customerData['初年度'] = convertExcelDateToString(customerData['初年度']);
                    }
                    
                    // フォルダ情報を追加
                    customerData['フォルダ名'] = folderName;
                    customerData['顧客コード'] = folderInfo.code;
                    customerData['顧客名'] = folderInfo.name;
                    
                    allCustomers.push(customerData);
                  }
                });
              }
            }
          }
        } catch (error) {
          console.error(`Error reading ${customerFilePath}:`, error);
        }
      }
    }

    console.log(`Loaded ${allCustomers.length} customers from ${customerFolders.length} folders`);
    return allCustomers;
    
  } catch (error) {
    console.error('Error reading customers folder:', error);
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
  if (match) {
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
    res.status(500).json({ error: 'Failed to read Excel', detail: String(err) });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  // eslint-disable-next-line no-console
  console.log(`Server listening on http://localhost:${PORT}`);
});
