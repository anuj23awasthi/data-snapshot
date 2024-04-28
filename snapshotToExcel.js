const puppeteer = require('puppeteer');
const Tesseract = require('tesseract.js');
const fs = require('fs');
const ExcelJS = require('exceljs');
const Jimp = require('jimp');

(async () => {
  try {
    const browser = await puppeteer.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    const page = await browser.newPage();
    await page.setViewport({ width: 1366, height: 768 });

    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.55 Safari/537.36');

    const url = 'https://www.nseindia.com/get-quotes/equity?symbol=TATAMOTORS';
    await page.goto(url);

    await page.waitForTimeout(30000);

    const screenshotPath = 'screenshot.png';
    await page.screenshot({ path: screenshotPath });

    await browser.close();

    const image = await Jimp.read(screenshotPath);

    // Define crop coordinates and dimensions for each section
    const cropCoordinates = [
      { x: 4, y: 100, width: 400, height: 200 }, // Customize these values
      { x: 14, y: 200, width: 400, height: 200 }, // Customize these values
      { x: 24, y: 300, width: 400, height: 200 },
      { x: 34, y: 400, width: 400, height: 200 },
    ];

    for (let i = 0; i < cropCoordinates.length; i++) {
      const coords = cropCoordinates[i];
      const croppedImage = await image.clone().crop(coords.x, coords.y, coords.width, coords.height);
      const croppedImagePath = `cropped_image_${i}.png`;
      await croppedImage.writeAsync(croppedImagePath);
      console.log(`Cropped image ${i} saved.`);

      const { data: { text } } = await Tesseract.recognize(croppedImagePath, 'eng', { logger: m => console.log(m) });
      console.log(`Extracted Text from Cropped Image ${i}:`, text);

      if (text) {
        const extractedData = processTradeInformation(text);

        const jsonPath = `trade_info_data_${i}.json`;
        const jsonData = JSON.stringify(extractedData, null, 2);
        fs.writeFileSync(jsonPath, jsonData, 'utf-8');
        console.log(`Trade information data from Cropped Image ${i} saved to`, jsonPath);

        const excelPath = `trade_info_data_${i}.xlsx`;
        saveToExcel(extractedData, excelPath);
        console.log(`Trade information data from Cropped Image ${i} saved to`, excelPath);
      } else {
        console.log(`No text extracted from Cropped Image ${i}. OCR process did not yield any results.`);
      }
    }
  } catch (error) {
    console.error('Error:', error);
  }
})();

function processTradeInformation(text) {
  const lines = text.split('\n');
  const data = [];

  let isTradeInfoSection = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    if (line === 'Buyer / Seller' || line === 'Price (in Rs.)' || line === 'Total Quantity') {
      isTradeInfoSection = true;
      continue;
    }

    if (isTradeInfoSection && line === '') {
      isTradeInfoSection = false;
      break;
    }

    if (isTradeInfoSection) {
      const values = line.split(/\s+/);
      if (values.length >= 3) {
        const entry = {
          Parameter: values[0],
          Value: values[1],
          Change: values[2],
        };
        data.push(entry);
      }
    }
  }

  return data;
}

function saveToExcel(data, filePath) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Trade Information Data');
  
  worksheet.addRow(Object.keys(data[0]));

  for (const entry of data) {
    worksheet.addRow(Object.values(entry));
  }

  workbook.xlsx.writeFile(filePath);
}

