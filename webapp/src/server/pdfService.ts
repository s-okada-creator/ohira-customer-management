import puppeteer from 'puppeteer';

interface CustomerData {
  '自宅郵便番号': string;
  '自宅住所1': string;
  '自宅住所2': string;
  '名　前': string;
}

interface SenderInfo {
  postalCode: string;
  lines: string[];
}

/**
 * はがき宛名PDFを生成
 */
export async function generateHagakiPDF(
  customers: CustomerData[],
  senderInfo: SenderInfo
): Promise<Buffer> {
  const html = generateHagakiHTML(customers, senderInfo);

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });

  try {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0' });

    const pdfBuffer = await page.pdf({
      format: 'A6',
      width: '100mm',
      height: '148mm',
      printBackground: true,
      margin: { top: 0, right: 0, bottom: 0, left: 0 },
    });

    return pdfBuffer as Buffer;
  } finally {
    await browser.close();
  }
}

/**
 * はがき宛名HTMLを生成
 */
function generateHagakiHTML(customers: CustomerData[], senderInfo: SenderInfo): string {
  const sheets = customers.map((customer) => generateSingleSheet(customer, senderInfo)).join('');

  return `
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>ハガキ宛名面</title>
  <style>
    @page {
      size: 100mm 148mm;
      margin: 0;
    }
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }
    body {
      width: 100mm;
      height: 148mm;
      font-family: "Noto Sans JP", "Hiragino Kaku Gothic ProN", "Yu Gothic", sans-serif;
      background: white;
      overflow: hidden;
    }
    .sheet {
      position: relative;
      width: 100mm;
      height: 148mm;
      background: white;
      page-break-after: always;
      overflow: hidden;
    }
    
    /* 切手貼付欄 */
    .stamp-box {
      position: absolute;
      top: 8mm;
      left: 8mm;
      width: 20mm;
      height: 25mm;
      border: 1.5px dashed #333;
      background: transparent;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 9pt;
      color: #333;
      writing-mode: vertical-rl;
      text-orientation: upright;
      letter-spacing: 1px;
    }
    
    /* 受取人郵便番号欄 */
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
    
    /* 宛名エリア */
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
    
    /* 差出人エリア */
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
    .sender-postal {
      display: flex;
      gap: 1mm;
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
  </style>
</head>
<body>
  ${sheets}
</body>
</html>
  `;
}

/**
 * 1枚のはがきHTMLを生成
 */
function generateSingleSheet(customer: CustomerData, senderInfo: SenderInfo): string {
  const fmt = (v: unknown) => (v == null ? '' : String(v));

  // 受取人情報
  const rawZip = fmt(customer['自宅郵便番号']);
  const address1 = fmt(customer['自宅住所1']);
  const address2 = fmt(customer['自宅住所2']);
  const name = fmt(customer['名　前']);

  // 文字変換
  const verticalAddress1 = convertToVerticalText(address1);
  const verticalAddress2 = convertToVerticalText(address2);
  const verticalName = convertToVerticalText(name);

  // 郵便番号処理
  const recipientPostal = renderPostalDigits(rawZip, false);
  const senderPostal = renderPostalDigits(senderInfo.postalCode, true);

  // 住所行
  const addressLines = [verticalAddress1, verticalAddress2]
    .filter(Boolean)
    .map((line) => `<span>${line}</span>`)
    .join('');

  // 差出人情報（右から：住所、名前、電話番号）
  const senderAddress = convertToVerticalText(senderInfo.lines[0] || '');
  const senderName = convertToVerticalText(senderInfo.lines[1] || '');
  const senderContact = convertToVerticalText(senderInfo.lines[2] || '');

  const senderInfoHTML = [senderContact, senderName, senderAddress]
    .filter(Boolean)
    .map((line) => `<div class="sender-line">${line}</div>`)
    .join('');

  return `
    <div class="sheet">
      <div class="stamp-box">切手貼付位置</div>
      <div class="recipient-postal">${recipientPostal}</div>
      <div class="recipient-address">${addressLines}</div>
      <div class="recipient-name">${verticalName}様</div>
      <div class="sender-block">
        <div class="sender-info">${senderInfoHTML}</div>
        <div class="sender-postal">${senderPostal}</div>
      </div>
    </div>
  `;
}

/**
 * 縦書き用にテキストを変換
 */
function convertToVerticalText(text: string): string {
  if (!text) return '';

  let converted = text.replace(/[!-~]/g, (ch) => {
    return String.fromCharCode(ch.charCodeAt(0) + 0xfee0);
  });

  converted = converted.replace(/ /g, '　');
  converted = converted.replace(/[-－]/g, 'ー');

  return converted;
}

/**
 * 郵便番号の数字枠を生成
 */
function renderPostalDigits(zipCode: string, isSender: boolean): string {
  const digits = zipCode.replace(/[^0-9]/g, '').slice(0, 7).padEnd(7, ' ');

  if (isSender) {
    // 差出人郵便番号（点線枠）
    return digits
      .split('')
      .map((d, i) => {
        if (i === 3) {
          return `<div style="width: 2mm;"></div><span class="sender-postal-box">${d.trim() ? d : ''}</span>`;
        }
        return `<span class="sender-postal-box">${d.trim() ? d : ''}</span>`;
      })
      .join('');
  } else {
    // 受取人郵便番号（赤枠、ハイフン付き）
    return digits
      .split('')
      .map((d, i) => {
        if (i === 3) {
          return `<div class="postal-separator">ー</div><span class="postal-box">${d.trim() ? d : ''}</span>`;
        }
        return `<span class="postal-box">${d.trim() ? d : ''}</span>`;
      })
      .join('');
  }
}

