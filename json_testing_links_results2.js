const puppeteer = require('puppeteer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const slugify = require('slugify');
const html2plaintext = require('html2plaintext');

slugify.extend({ ':': '-', '/': '_' });

const esgTerms = require('./esg_terms.json');
const banksToCrawl = require('./banks-to-crawl.json');
const linksFilePath = path.join(__dirname, 'links.xlsx');
const resultsFilePath = path.join(__dirname, 'results.xlsx');

// Step 1: Crawl the home pages of banks and save the links to an Excel file
async function crawlHomePages() {
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();
  const allLinks = [];

  for (const baseUrl of banksToCrawl) {
    try {
      await page.goto(baseUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await page.evaluate(() => new Promise(resolve => setTimeout(resolve, 5000))); // Wait for 5 seconds

      const links = await page.evaluate(baseUrl => {
        return Array.from(document.querySelectorAll('a'))
          .map(link => link.href)
          .filter(href => href && href.startsWith(baseUrl));
      }, baseUrl);

      allLinks.push(...links);
    } catch (error) {
      console.error(`Failed to crawl ${baseUrl}:`, error);
    }
  }

  await browser.close();
  writeLinksToExcel(allLinks);
}

function writeLinksToExcel(links) {
  const headers = ['Links'];
  const data = links.map(link => [link]);

  const worksheet = xlsx.utils.aoa_to_sheet([headers, ...data]);
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Links');

  xlsx.writeFile(workbook, linksFilePath);

  console.log(`Links written to ${linksFilePath}`);
}

// Step 2: Read the links from the Excel file
function readLinksFromExcel() {
  const workbook = xlsx.readFile(linksFilePath);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const links = xlsx.utils.sheet_to_json(worksheet, { header: 1 }).slice(1).map(row => row[0]);
  return links;
}

// Step 3: Crawl each link and search for ESG terms
async function crawlLinksAndSearchESGTerms() {
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();
  const links = readLinksFromExcel();
  const results = [];

  for (const link of links) {
    try {
      await page.goto(link, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await page.evaluate(() => new Promise(resolve => setTimeout(resolve, 5000))); // Wait for 5 seconds

      const content = await page.evaluate(() => {
        const body = document.querySelector('body');
        return body ? body.innerText : '';
      });

      const plainText = html2plaintext(content).replace(/\s+/g, ' ').trim();
      const sluggedUrl = slugify(link);
      const fileName = `${sluggedUrl}.txt`;

      fs.writeFileSync(path.join(__dirname, 'parsing_test', fileName), plainText);

      const result = searchWords(plainText, esgTerms);
      results.push({ website: link, ...result });
    } catch (error) {
      console.error(`Failed to crawl ${link}:`, error);
    }
  }

  await browser.close();
  writeResultsToExcel(results);
}

function searchWords(text, words) {
  const result = {};
  words.forEach(word => {
    const regex = new RegExp(`\\b${word}\\b`, 'gi');
    const matches = text.match(regex);
    result[word] = matches ? matches.length : 0;
  });
  return result;
}

// Step 4: Save the results to another Excel file
function writeResultsToExcel(results) {
  const headers = ['Website', ...esgTerms];
  const data = results.map(result => [result.website, ...esgTerms.map(word => result[word] || 0)]);

  const worksheet = xlsx.utils.aoa_to_sheet([headers, ...data]);
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Results');

  xlsx.writeFile(workbook, resultsFilePath);

  console.log(`Results written to ${resultsFilePath}`);
}

async function main() {
  await crawlHomePages();
  await crawlLinksAndSearchESGTerms();
}

main().catch(error => {
  console.error('Error:', error);
});
