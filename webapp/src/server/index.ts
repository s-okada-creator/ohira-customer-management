import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';
import { getCustomerSheetValues } from './googleSheets.js';
import puppeteer from 'puppeteer';
import puppeteerCore from 'puppeteer-core';
import chromium from '@sparticuz/chromium';
import { PDFDocument } from 'pdf-lib';
import multer from 'multer';
import { put, list, del } from '@vercel/blob';

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

// はがき裏面画像のアップロード設定（Vercel Blob使用）
// Blobストアのプレフィックス（同じプロジェクト内で他用途と区別）
const HAGAKI_BLOB_PREFIX = 'hagaki-backs/';

// multerはメモリストレージ（Blobへ転送するためディスクには書かない）
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }, // 10MB
  fileFilter: (_req, file, cb) => {
    const allowed = ['.png', '.jpg', '.jpeg', '.pdf'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowed.includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('PNG, JPG, PDFのみアップロード可能です'));
    }
  }
});

function hasBlobToken(): boolean {
  return !!process.env.BLOB_READ_WRITE_TOKEN;
}

app.get('/health', (_req, res) => {
  res.json({ ok: true, blob: hasBlobToken() });
});

// はがき裏面画像一覧取得
app.get('/api/hagaki-backs', async (_req, res) => {
  try {
    const files: { id: string; name: string; url: string; isDefault: boolean }[] = [];

    // デフォルト画像
    files.push({
      id: 'default',
      name: 'お得で安心♪（デフォルト）',
      url: '/otoku-anshin-back.png',
      isDefault: true
    });

    if (hasBlobToken()) {
      // Vercel Blobから一覧取得
      const { blobs } = await list({ prefix: HAGAKI_BLOB_PREFIX });
      for (const blob of blobs) {
        // pathnameは "hagaki-backs/xxx.png" の形式
        const filename = blob.pathname.substring(HAGAKI_BLOB_PREFIX.length);
        if (!filename) continue;
        files.push({
          id: encodeURIComponent(blob.url),
          name: filename,
          url: blob.url,
          isDefault: false
        });
      }
    }

    res.json({ files });
  } catch (error) {
    console.error('裏面画像一覧の取得エラー:', error);
    res.status(500).json({ error: '画像一覧の取得に失敗しました', detail: String(error) });
  }
});

// はがき裏面画像アップロード
app.post('/api/hagaki-backs/upload', upload.single('image'), async (req: any, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'ファイルが選択されていません' });
    }
    if (!hasBlobToken()) {
      return res.status(500).json({ error: 'Blobストレージが設定されていません（BLOB_READ_WRITE_TOKEN未設定）' });
    }

    const file = req.file;
    const ext = path.extname(file.originalname);
    // ファイル名衝突を避けるためタイムスタンプ + ランダムを付与
    const safeName = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}${ext}`;
    const pathname = `${HAGAKI_BLOB_PREFIX}${safeName}`;

    const blob = await put(pathname, file.buffer, {
      access: 'public',
      contentType: file.mimetype,
      addRandomSuffix: false
    });

    res.json({
      id: encodeURIComponent(blob.url),
      name: file.originalname,
      url: blob.url,
      isDefault: false
    });
  } catch (error) {
    console.error('裏面画像アップロードエラー:', error);
    res.status(500).json({ error: 'アップロードに失敗しました', detail: String(error) });
  }
});

// はがき裏面画像削除（URLをクエリパラメータで受け取る）
app.delete('/api/hagaki-backs', async (req, res) => {
  try {
    const blobUrl = typeof req.query.url === 'string' ? req.query.url : '';
    if (!blobUrl) {
      return res.status(400).json({ error: '削除対象URLが指定されていません' });
    }
    if (!hasBlobToken()) {
      return res.status(500).json({ error: 'Blobストレージが設定されていません' });
    }
    // 安全のため、自分のhagaki-backsプレフィックスのもののみ削除許可
    if (!blobUrl.includes(HAGAKI_BLOB_PREFIX)) {
      return res.status(400).json({ error: '削除対象が不正です' });
    }

    await del(blobUrl);
    res.json({ ok: true });
  } catch (error) {
    console.error('裏面画像削除エラー:', error);
    res.status(500).json({ error: '削除に失敗しました', detail: String(error) });
  }
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

// はがきPDF生成エンドポイント
app.post('/api/generate-hagaki-pdf', async (req, res) => {
  try {
    const { customerIds, backImageUrl } = req.body;

    if (!customerIds || !Array.isArray(customerIds)) {
      return res.status(400).json({ error: '顧客IDが指定されていません' });
    }

    console.log(`はがきPDF生成開始: ${customerIds.length}件の顧客, backImageUrl=${backImageUrl || '(default)'}`);

    // 全顧客データを取得
    const allCustomers = await readCustomers();

    // 指定された顧客のみをフィルタリング
    const targetCustomers = customerIds[0] === 'all'
      ? allCustomers
      : allCustomers.filter(c => customerIds.includes(c['顧客番号']));

    if (targetCustomers.length === 0) {
      return res.status(400).json({ error: '対象顧客が見つかりません' });
    }

    // はがき宛名PDFを生成
    const addressPdfBytes = await generateAddressPdf(targetCustomers);

    // 裏面画像を取得（カスタム or デフォルト）
    const backPdfBytes = await loadBackPdfBytes(backImageUrl);
    
    // 表面と裏面を結合
    const finalPdf = await combinePdfs(addressPdfBytes, backPdfBytes, targetCustomers.length);
    
    console.log(`はがきPDF生成完了: ${targetCustomers.length}件`);
    
    // PDFをレスポンスとして返す
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="hagaki_${new Date().toISOString().slice(0, 10)}.pdf"`);
    res.send(Buffer.from(finalPdf));
    
  } catch (error) {
    console.error('PDF生成エラー:', error);
    res.status(500).json({ error: 'PDF生成に失敗しました', detail: String(error) });
  }
});

/**
 * 裏面PDFバイト列を取得
 * - backImageUrl 未指定 or 'default' の場合: デフォルトのお得で安心PDF
 * - URL指定の場合: そのURLからフェッチし、画像ならPDFに変換
 */
async function loadBackPdfBytes(backImageUrl?: string): Promise<Buffer> {
  const isCustom = backImageUrl && backImageUrl !== '/otoku-anshin-back.png' && backImageUrl !== 'default';

  if (!isCustom) {
    const backPdfPath = path.resolve(process.cwd(), 'src', 'public', 'お得で安心♪.pdf');
    if (!fs.existsSync(backPdfPath)) {
      throw new Error(`デフォルト裏面PDFが見つかりません: ${backPdfPath}`);
    }
    return fs.readFileSync(backPdfPath);
  }

  // 外部URL（Vercel Blobなど）からフェッチ
  const fetchUrl = backImageUrl!.startsWith('http')
    ? backImageUrl!
    : `${process.env.VERCEL_URL ? 'https://' + process.env.VERCEL_URL : 'http://localhost:' + (process.env.PORT || 3000)}${backImageUrl}`;

  const response = await fetch(fetchUrl);
  if (!response.ok) {
    throw new Error(`裏面画像の取得に失敗しました (${response.status}): ${fetchUrl}`);
  }
  const arrayBuffer = await response.arrayBuffer();
  const buffer = Buffer.from(arrayBuffer);

  const contentType = (response.headers.get('content-type') || '').toLowerCase();
  const urlLower = fetchUrl.toLowerCase();
  const isPdf = contentType.includes('pdf') || urlLower.endsWith('.pdf');

  if (isPdf) {
    return buffer;
  }

  // 画像 → ハガキサイズPDFに変換
  const ext = urlLower.endsWith('.png') ? '.png' : '.jpg';
  return Buffer.from(await convertImageToPdf(buffer, ext));
}

/**
 * 画像バッファをはがきサイズPDFに変換
 */
async function convertImageToPdf(imageData: Buffer, ext: string): Promise<Uint8Array> {
  const isVercel = process.env.VERCEL === '1' || process.env.NODE_ENV === 'production';
  let browser;
  if (isVercel) {
    browser = await puppeteerCore.launch({
      args: chromium.args,
      executablePath: await chromium.executablePath(),
      headless: true,
    });
  } else {
    browser = await puppeteer.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
  }
  try {
    const page = await browser.newPage();
    const mimeType = ext === '.png' ? 'image/png' : 'image/jpeg';
    const base64 = imageData.toString('base64');
    const html = `<!DOCTYPE html><html><head><style>
      @page{size:100mm 148mm;margin:0;}
      body{margin:0;padding:0;}
      img{width:100mm;height:148mm;object-fit:contain;display:block;}
    </style></head><body>
      <img src="data:${mimeType};base64,${base64}" />
    </body></html>`;
    await page.setContent(html, { waitUntil: 'networkidle0' });
    const pdfBuffer = await page.pdf({
      width: '100mm',
      height: '148mm',
      printBackground: true,
      margin: { top: 0, right: 0, bottom: 0, left: 0 }
    });
    return new Uint8Array(pdfBuffer);
  } finally {
    await browser.close();
  }
}

/**
 * はがき宛名PDFを生成（Puppeteerを使用）
 */
async function generateAddressPdf(customers: CustomerRow[]): Promise<Uint8Array> {
  // Vercel環境かどうかを判定
  const isVercel = process.env.VERCEL === '1' || process.env.NODE_ENV === 'production';
  
  let browser;
  
  if (isVercel) {
    // Vercel環境: puppeteer-core + chromium を使用
    browser = await puppeteerCore.launch({
      args: chromium.args,
      executablePath: await chromium.executablePath(),
      headless: true,
    });
  } else {
    // ローカル環境: 通常のpuppeteerを使用
    browser = await puppeteer.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
  }
  
  try {
    const page = await browser.newPage();
    
    // はがき宛名HTMLを生成
    const html = generateAddressHtml(customers);
    
    await page.setContent(html, { waitUntil: 'networkidle0' });
    
    // PDFを生成（はがきサイズ: 100mm x 148mm）
    const pdfBuffer = await page.pdf({
      width: '100mm',
      height: '148mm',
      printBackground: true,
      margin: { top: 0, right: 0, bottom: 0, left: 0 }
    });
    
    return new Uint8Array(pdfBuffer);
  } finally {
    await browser.close();
  }
}

/**
 * 表面（宛名）と裏面（お得で安心PDF）を結合
 */
async function combinePdfs(addressPdfBytes: Uint8Array, backPdfBytes: Buffer, customerCount: number): Promise<Uint8Array> {
  // 新しいPDFドキュメントを作成
  const finalPdf = await PDFDocument.create();
  
  // 宛名PDFを読み込み
  const addressPdf = await PDFDocument.load(addressPdfBytes);
  const addressPages = await finalPdf.copyPages(addressPdf, addressPdf.getPageIndices());
  
  // お得で安心PDFを読み込み（裏面用）
  const backPdf = await PDFDocument.load(backPdfBytes);
  // 裏面は最初の1ページのみ使用
  const [backPage] = await finalPdf.copyPages(backPdf, [0]);
  
  // 各顧客に対して、表面（宛名）→ 裏面（お得で安心の1ページ目）の順で追加
  for (let i = 0; i < customerCount; i++) {
    // 表面（宛名）を追加
    if (i < addressPages.length) {
      finalPdf.addPage(addressPages[i]);
    }
    
    // 裏面（お得で安心PDFの1ページ目）を追加
    // 各顧客ごとに新しいコピーを作成
    if (i === 0) {
      // 最初の顧客には既にコピーしたページを使用
      finalPdf.addPage(backPage);
    } else {
      // 2番目以降の顧客には新しいコピーを作成
      const [newBackPage] = await finalPdf.copyPages(backPdf, [0]);
      finalPdf.addPage(newBackPage);
    }
  }
  
  // PDFをバイト配列として保存
  const pdfBytes = await finalPdf.save();
  return pdfBytes;
}

/**
 * はがき宛名HTMLを生成
 */
function generateAddressHtml(customers: CustomerRow[]): string {
  const SENDER_INFO = {
    postalCode: '751-0804',
    lines: ['山口県下関市楠乃５丁目９', '大平自動車商会', 'TEL: 083-257-0101']
  };
  
  const fmt = (v: any) => v == null ? '' : String(v);
  
  const convertToVerticalText = (text: string): string => {
    if (!text) return '';
    let converted = text.replace(/[!-~]/g, (ch) => 
      String.fromCharCode(ch.charCodeAt(0) + 0xFEE0)
    );
    converted = converted.replace(/ /g, '　');
    converted = converted.replace(/[-－]/g, 'ー');
    return converted;
  };
  
  const formatPostalCode = (zipCode: string): string => {
    if (!zipCode) return '';
    const digits = zipCode.replace(/[^0-9]/g, '');
    if (digits.length !== 7) return zipCode;
    return `${digits.slice(0, 3)}-${digits.slice(3)}`;
  };
  
  const renderPostalDigits = (originalZip: string, isSender = false): string => {
    const digits = originalZip.replace(/[^0-9]/g, '').slice(0, 7);
    const padded = digits.padEnd(7, ' ');
    
    if (isSender) {
      return padded.split('').map((d, i) => {
        if (i === 3) {
          return `<div style="width: 2mm;"></div><span class="sender-postal-box">${d.trim() ? d : ''}</span>`;
        }
        return `<span class="sender-postal-box">${d.trim() ? d : ''}</span>`;
      }).join('');
    } else {
      return padded.split('').map((d, i) => {
        if (i === 3) {
          return `<div class="postal-separator">ー</div><span class="postal-box">${d.trim() ? d : ''}</span>`;
        }
        return `<span class="postal-box">${d.trim() ? d : ''}</span>`;
      }).join('');
    }
  };
  
  const splitAddress = (address: string): { line1: string; line2: string } => {
    if (!address) return { line1: '', line2: '' };
    const spaceMatch = address.match(/^([^\s　]+)[\s　]+(.+)$/);
    if (spaceMatch && spaceMatch[1] && spaceMatch[2]) {
      return { line1: spaceMatch[1], line2: spaceMatch[2] };
    }
    if (address.length >= 40) {
      const midPoint = Math.floor(address.length / 2);
      return { line1: address.slice(0, midPoint), line2: address.slice(midPoint) };
    }
    return { line1: address, line2: '' };
  };
  
  const renderAddressCard = (row: CustomerRow): string => {
    const rawZip = fmt(row['自宅郵便番号']);
    const address1 = fmt(row['自宅住所1']);
    const address2 = fmt(row['自宅住所2']);
    const name = fmt(row['名　前']) || fmt(row['顧客名']) || fmt(row['顧客名（フォルダ）']);

    const formattedZip = formatPostalCode(rawZip);
    const zipDigits = formattedZip.replace(/[^0-9]/g, '');

    let addressLine1, addressLine2;
    if (address1 && address2) {
      addressLine1 = address1;
      addressLine2 = address2;
    } else if (address1 && !address2) {
      const split = splitAddress(address1);
      addressLine1 = split.line1;
      addressLine2 = split.line2;
    } else {
      addressLine1 = '';
      addressLine2 = '';
    }

    const verticalAddress1 = convertToVerticalText(addressLine1);
    const verticalAddress2 = convertToVerticalText(addressLine2);
    const verticalName = convertToVerticalText(name);

    const postalElement = zipDigits ? `<div class="recipient-postal">${renderPostalDigits(zipDigits)}</div>` : '';
    const senderZip = formatPostalCode(SENDER_INFO.postalCode);
    const senderZipDigits = senderZip.replace(/[^0-9]/g, '');
    const senderPostal = senderZipDigits ? `<div class="sender-postal">${renderPostalDigits(senderZipDigits, true)}</div>` : '';
    
    const addressLines = [verticalAddress1, verticalAddress2]
      .filter(Boolean)
      .map(line => `<span>${line}</span>`)
      .join('');
    
    const senderAddress = convertToVerticalText(SENDER_INFO.lines[0] || '');
    const senderName = convertToVerticalText(SENDER_INFO.lines[1] || '');
    const senderContact = convertToVerticalText(SENDER_INFO.lines[2] || '');
    
    const senderInfo = [senderContact, senderName, senderAddress]
      .filter(Boolean)
      .map(line => `<div class="sender-line">${line}</div>`)
      .join('');
    
    const nameWithHonorific = verticalName ? `<div class="recipient-name">${verticalName}様</div>` : '';
    const stampBox = `<div class="stamp-box">切手貼付位置</div>`;

    return `
      <div class="sheet">
        ${stampBox}
        ${postalElement}
        <div class="recipient-address">${addressLines}</div>
        ${nameWithHonorific}
        <div class="sender-block">
          <div class="sender-info">${senderInfo}</div>
          ${senderPostal}
        </div>
      </div>
    `;
  };
  
  const cards = customers.map(renderAddressCard).join('');
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>はがき宛名</title>
      <style>
        @page {
          size: 100mm 148mm;
          margin: 0;
        }
        body {
          margin: 0;
          padding: 0;
          font-family: 'Noto Sans JP', 'Hiragino Sans', 'Hiragino Kaku Gothic ProN', 'Yu Gothic', '游ゴシック', sans-serif;
        }
        .sheet {
          width: 100mm;
          height: 148mm;
          background: white;
          page-break-after: always;
          position: relative;
          box-sizing: border-box;
        }
        .stamp-box {
          position: absolute;
          top: 8mm;
          left: 8mm;
          width: 20mm;
          height: 25mm;
          border: 1.5px dashed #333;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 9pt;
          color: #333;
          writing-mode: vertical-rl;
          text-orientation: upright;
          letter-spacing: 1px;
        }
        .recipient-postal {
          position: absolute;
          top: 10mm;
          right: 8mm;
          display: flex;
          gap: 1mm;
        }
        .postal-box {
          width: 5.5mm;
          height: 7mm;
          border: 1.5px solid #D32F2F;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 14pt;
          font-weight: bold;
          color: black;
          background: white;
          font-family: "Courier New", monospace;
        }
        .postal-separator {
          width: 2mm;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 16pt;
          color: #D32F2F;
        }
        .sender-postal {
          display: flex;
          gap: 1mm;
          margin-bottom: 2mm;
        }
        .sender-postal-box {
          width: 4mm;
          height: 5mm;
          border: 1px dashed #D32F2F;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 9pt;
          color: black;
          background: white;
          font-family: "Courier New", monospace;
        }
        .recipient-address {
          position: absolute;
          top: 25mm;
          right: 8mm;
          font-size: 12pt;
          color: black;
          writing-mode: vertical-rl;
          text-orientation: upright;
          line-height: 1.6;
          letter-spacing: 1px;
        }
        .recipient-address span {
          display: inline-block;
          margin-left: 5mm;
        }
        .recipient-name {
          position: absolute;
          top: 35mm;
          right: 45mm;
          font-size: 24pt;
          font-weight: bold;
          color: black;
          writing-mode: vertical-rl;
          text-orientation: upright;
          line-height: 1.4;
          letter-spacing: 3px;
        }
        .sender-block {
          position: absolute;
          bottom: 8mm;
          left: 8mm;
          width: 70mm;
        }
        .sender-info {
          display: flex;
          flex-direction: row;
          gap: 3mm;
          margin-bottom: 4mm;
        }
        .sender-line {
          font-size: 10pt;
          color: black;
          writing-mode: vertical-rl;
          text-orientation: upright;
          letter-spacing: 1px;
          line-height: 1.4;
        }
      </style>
    </head>
    <body>
      ${cards}
    </body>
    </html>
  `;
}

// Vercel環境では app.listen() は不要
if (process.env.NODE_ENV !== 'production') {
  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () => {
    // eslint-disable-next-line no-console
    console.log(`Server listening on http://localhost:${PORT}`);
  });
}

export default app;
