const puppeteer = require('puppeteer');
const excel = require('exceljs');

async function scrapeBooks() {
  const workbook = new excel.Workbook();
  await workbook.xlsx.readFile('dataIO/Input.xlsx');
  const worksheet = workbook.getWorksheet(1);
  const rows = worksheet.getRows();

  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  await page.goto('https://www.snapdeal.com/');

  for (let i = 1; i <= rows.length; i++) {
    const row = rows[i];
    const isbn = row.getCell(2).value;
    const title = row.getCell(3).value;

    await page.waitForSelector('#inputValEnter', { visible: true });
    await page.type('#inputValEnter', isbn);
    await page.click('button[type="submit"]');

    const searchResult = await page.waitForSelector('.search-result', {
      visible: true,
      timeout: 5000,
    });

    const resultText = await page.evaluate(
      (el) => el.textContent,
      searchResult
    );

    if (resultText.includes('No results found for')) {
      row.getCell(4).value = 'No';
      continue;
    }

    const bookTitles = await page.$$('.product-title');

    let bestMatchTitle = '';
    let bestMatchPrice = Number.MAX_VALUE;

    for (const bookTitle of bookTitles) {
      const text = await page.evaluate((el) => el.textContent, bookTitle);
      if (text.toLowerCase().includes(title.toLowerCase())) {
        const bookPriceElement = await bookTitle.$('.product-price');
        if (bookPriceElement) {
          const bookPrice = await page.evaluate(
            (el) => Number(el.textContent.replace(/[^0-9.-]+/g, '')),
            bookPriceElement
          );
          if (bookPrice < bestMatchPrice) {
            bestMatchTitle = text;
            bestMatchPrice = bookPrice;
          }
        }
      }
    }

    if (bestMatchTitle) {
      row.getCell(4).value = 'Yes';
      row.getCell(5).value = bestMatchTitle;
      row.getCell(6).value = bestMatchPrice;
      await page.click(`a[title="${bestMatchTitle}"]`);
      await page.waitForSelector('#buy-button-id', { visible: true });
      const authorElement = await page.$('.publisher-name');
      const publisherElement = await page.$('.publisher-name span');
      const stockElement = await page.$('.inventory');
      const url = page.url();
      row.getCell(7).value = authorElement
        ? await page.evaluate((el) => el.textContent, authorElement)
        : '';
      row.getCell(8).value = publisherElement
        ? await page.evaluate((el) => el.textContent, publisherElement)
        : '';
      row.getCell(9).value = stockElement
        ? await page.evaluate((el) => el.textContent, stockElement)
        : '';
      row.getCell(10).value = url;
      await page.goBack();
    }
  }

  await workbook.xlsx.writeFile('Output.xlsx');
  await browser.close();
}

scrapeBooks();
