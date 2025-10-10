import { put } from '@vercel/blob';
import fs from 'fs';
import path from 'path';
import archiver from 'archiver';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function uploadCustomerData() {
  try {
    console.log('📦 Creating archive of customer data...');
    
    // Create a zip file of the customer folder
    const output = fs.createWriteStream('customer-data.zip');
    const archive = archiver('zip', { zlib: { level: 9 } });
    
    archive.pipe(output);
    
    // Add the customer folder to the archive
    const customerFolderPath = path.resolve(__dirname, '..', '☆顧客フォルダ');
    archive.directory(customerFolderPath, '☆顧客フォルダ');
    
    await archive.finalize();
    
    console.log('📤 Uploading to Vercel Blob Storage...');
    
    // Upload to Vercel Blob Storage
    const blob = await put('customer-data.zip', fs.readFileSync('customer-data.zip'), {
      access: 'public',
    });
    
    console.log('✅ Upload successful!');
    console.log('📎 Blob URL:', blob.url);
    
    // Clean up local zip file
    fs.unlinkSync('customer-data.zip');
    
    return blob.url;
  } catch (error) {
    console.error('❌ Upload failed:', error);
    throw error;
  }
}

// Run if called directly
if (import.meta.url === `file://${process.argv[1]}`) {
  uploadCustomerData()
    .then(url => {
      console.log('🎉 Customer data uploaded successfully!');
      console.log('🔗 URL:', url);
      process.exit(0);
    })
    .catch(error => {
      console.error('💥 Error:', error);
      process.exit(1);
    });
}

export { uploadCustomerData };
