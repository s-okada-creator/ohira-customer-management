import fs from 'fs';
import path from 'path';
import xlsx from 'xlsx';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function convertExcelDateToString(excelDate) {
  const excelEpoch = new Date(1900, 0, 1);
  const actualDate = new Date(excelEpoch.getTime() + (excelDate - 2) * 24 * 60 * 60 * 1000);
  
  const year = actualDate.getFullYear();
  const month = (actualDate.getMonth() + 1).toString().padStart(2, '0');
  const day = actualDate.getDate().toString().padStart(2, '0');
  
  return `${year}/${month}/${day}`;
}

function parseFolderName(folderName) {
  const match = folderName.match(/^(\d+)(.+)$/);
  if (match && match[1] && match[2]) {
    return {
      code: match[1],
      name: match[2].replace(/[()（）]/g, '')
    };
  }
  return { code: '', name: folderName };
}

async function convertCustomerData() {
  try {
    console.log('🔄 Converting customer data to JSON...');
    
    const customersFolder = path.resolve(__dirname, '..', '☆顧客フォルダ');
    
    if (!fs.existsSync(customersFolder)) {
      console.error('❌ Customers folder not found:', customersFolder);
      return;
    }
    
    const customerFolders = fs.readdirSync(customersFolder, { withFileTypes: true })
      .filter(dirent => dirent.isDirectory())
      .map(dirent => dirent.name);
    
    console.log(`📁 Found ${customerFolders.length} customer folders`);
    
    const allCustomers = [];
    
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
          
          if (!targetSheet && workbook.SheetNames.length > 0) {
            const firstSheetName = workbook.SheetNames[0];
            if (firstSheetName) {
              targetSheet = workbook.Sheets[firstSheetName];
            }
          }
          
          if (targetSheet) {
            const range = targetSheet['!ref'];
            if (range) {
              const rangeParts = range.split(':');
              const endCell = rangeParts[1];
              if (endCell) {
                const endRow = parseInt(endCell.replace(/[A-Z]/g, ''));
                
                if (endRow >= 2) {
                  const headerRow = xlsx.utils.sheet_to_json(targetSheet, { header: 1, range: 0 })[0];
                  const dataRows = xlsx.utils.sheet_to_json(targetSheet, { header: 1, range: 1 });
                  
                  dataRows.forEach(dataRow => {
                    if (dataRow.length > 0 && dataRow[1] && String(dataRow[1]).trim() !== '') {
                      const customerData = {};
                      
                      headerRow.forEach((header, index) => {
                        if (header && dataRow[index] !== undefined) {
                          customerData[header] = dataRow[index];
                        }
                      });
                      
                      const folderInfo = parseFolderName(folderName);
                      
                      // 日付フィールドを変換
                      if (customerData['次回車検満期日'] && typeof customerData['次回車検満期日'] === 'number') {
                        customerData['次回車検満期日'] = convertExcelDateToString(customerData['次回車検満期日']);
                      }
                      if (customerData['初年度'] && typeof customerData['初年度'] === 'number') {
                        customerData['初年度'] = convertExcelDateToString(customerData['初年度']);
                      }
                      
                      customerData['フォルダ名'] = folderName;
                      customerData['顧客コード'] = folderInfo.code;
                      customerData['顧客名'] = folderInfo.name;
                      
                      allCustomers.push(customerData);
                    }
                  });
                }
              }
            }
          }
        } catch (error) {
          console.error(`❌ Error reading ${customerFilePath}:`, error);
        }
      }
    }
    
    console.log(`✅ Converted ${allCustomers.length} customers to JSON`);
    
    // JSONファイルとして保存
    const outputPath = path.join(__dirname, 'src', 'data', 'customers.json');
    const outputDir = path.dirname(outputPath);
    
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }
    
    fs.writeFileSync(outputPath, JSON.stringify(allCustomers, null, 2));
    console.log(`💾 Saved to: ${outputPath}`);
    
    // ファイルサイズを確認
    const stats = fs.statSync(outputPath);
    const fileSizeInMB = stats.size / (1024 * 1024);
    console.log(`📊 File size: ${fileSizeInMB.toFixed(2)} MB`);
    
    if (fileSizeInMB > 50) {
      console.log('⚠️  Warning: File size exceeds Vercel limit (50MB)');
    } else {
      console.log('✅ File size is within Vercel limit');
    }
    
  } catch (error) {
    console.error('💥 Error converting data:', error);
  }
}

convertCustomerData();
