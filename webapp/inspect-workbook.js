import XLSX from 'xlsx';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 統合Excelブックのパス
const workbookPath = path.resolve(__dirname, '..', '顧客データ統合_全シート_20251003_194945.xlsx');

console.log('=== 統合Excelブック構造調査 ===');
console.log('ファイルパス:', workbookPath);

try {
  const workbook = XLSX.readFile(workbookPath);
  const sheetNames = workbook.SheetNames;
  
  console.log('\n=== シート一覧 ===');
  sheetNames.forEach((name, index) => {
    console.log(`${index + 1}. ${name}`);
  });
  
  // 目次シートを除外
  const dataSheets = sheetNames.filter(name => name !== '目次');
  console.log(`\nデータシート数: ${dataSheets.length} (目次除く)`);
  
  // 最初の数シートの構造を調査
  console.log('\n=== シート構造調査 ===');
  const sampleSheets = dataSheets.slice(0, 3);
  
  sampleSheets.forEach((sheetName) => {
    console.log(`\n--- ${sheetName} ---`);
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (data.length > 0) {
      const headers = data[0];
      console.log('列数:', headers.length);
      console.log('ヘッダー:', headers);
      
      if (data.length > 1) {
        console.log('サンプルデータ行:', data[1]);
      }

      console.log('\n------ 先頭〜5行 ------');
      data.slice(0, 5).forEach((row, idx) => {
        console.log(`#${idx}`, row);
      });
    } else {
      console.log('データなし');
    }
  });
  
  // 全シートの行数を集計
  console.log('\n=== 全シート行数集計 ===');
  let totalRows = 0;
  dataSheets.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    const rowCount = data.length - 1; // ヘッダー行を除く
    totalRows += rowCount;
    console.log(`${sheetName}: ${rowCount}行`);
  });
  
  console.log(`\n総データ行数: ${totalRows}行`);
  
} catch (error) {
  console.error('エラー:', error.message);
}
