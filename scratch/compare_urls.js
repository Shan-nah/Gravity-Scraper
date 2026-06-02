require('dotenv').config();
const axios = require('axios');
const cheerio = require('cheerio');
const H = {'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36','Accept':'text/html'};
const clean = s => (s||'').replace(/\s+/g,' ').trim();

async function analyzeUrl(label, url) {
  console.log(`\n=== ${label} ===`);
  console.log('URL:', url);
  const r = await axios.get(url, {headers:H, timeout:20000});
  const $ = cheerio.load(r.data);
  
  // Count key elements
  console.log('m-mainTR:', $('div.m-mainTR').length);
  console.log('m-r-government-tenders:', $('p.m-r-government-tenders').length);
  
  // Examine first m-mainTR structure
  const first = $('div.m-mainTR').first();
  const rows = first.children('div.row');
  console.log('Children div.row count in first m-mainTR:', rows.length);
  
  rows.each((i, row) => {
    const $row = $(row);
    const links = $row.find('a[href]').map((_,a) => $(a).attr('href')).get();
    console.log(`  div.row[${i}]: classes="${$row.attr('class')}", links=${JSON.stringify(links.slice(0,2))}`);
  });
  
  // Get viewLink using current code
  const row2 = first.children('div.row').eq(1);
  const viewLink = row2.find('a').first().attr('href') || null;
  console.log('viewLink from row[1]:', viewLink);
  
  // Try row[0] too
  const row0 = first.children('div.row').eq(0);
  const viewLink0 = row0.find('a').first().attr('href') || null;
  console.log('viewLink from row[0]:', viewLink0);
  
  // Show raw HTML of first m-mainTR
  console.log('First m-mainTR HTML:', first.html()?.replace(/\s+/g,' ').slice(0,600));
}

(async () => {
  await analyzeUrl('WORKING', 'https://www.tenderdetail.com/dailytenders/51168758/66998cc5-1b3c-4664-be2e-df09f4bf3c1d');
  await analyzeUrl('BROKEN', 'https://www.tenderdetail.com/dailytenders/51188939/226e8ac7-3ff2-4dc0-94b7-31dd4de4d76f');
})().catch(e => console.error('ERROR:', e.message));
