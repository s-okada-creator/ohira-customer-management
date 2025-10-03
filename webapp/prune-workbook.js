import XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const workbookPath = path.resolve(
  __dirname,
  '..',
  '顧客データ統合_全シート_20251003_194945.xlsx'
);

if (!fs.existsSync(workbookPath)) {
  console.error('ファイルが見つかりません:', workbookPath);
  process.exit(1);
}

console.log('対象ファイル:', workbookPath);

const workbook = XLSX.readFile(workbookPath);
const originalSheets = workbook.SheetNames;

const keepSheets = originalSheets.filter((name) => name.endsWith('_顧客名簿'));
const removedSheets = originalSheets.filter((name) => !keepSheets.includes(name));

console.log('総シート数:', originalSheets.length);
console.log('保持するシート数:', keepSheets.length);
console.log('削除対象シート数:', removedSheets.length);

const backupPath = workbookPath.replace(
  /\.xlsx$/i,
  `_backup_${new Date().toISOString().replace(/[:.]/g, '-')}.xlsx`
);

fs.copyFileSync(workbookPath, backupPath);
console.log('バックアップ作成:', backupPath);

const newWorkbook = XLSX.utils.book_new();

for (const sheetName of keepSheets) {
  const sheet = workbook.Sheets[sheetName];
  if (sheet) {
    XLSX.utils.book_append_sheet(newWorkbook, sheet, sheetName);
  }
}

XLSX.writeFile(newWorkbook, workbookPath, { bookType: 'xlsx', bookSST: true });

console.log('顧客名簿シートのみ残したファイルに更新しました。');

