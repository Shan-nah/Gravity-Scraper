require('dotenv').config();
const axios = require('axios');
const cheerio = require('cheerio');
const H = {'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36','Accept':'text/html'};
const clean = s => (s||'').replace(/\s+/g,' ').trim();

async function getLinks(url) {
  const r = await axios.get(url, {headers:H, timeout:20000});
  const $ = cheerio.load(r.data);
  const links = [];
  $('div.m-mainTR').each((_, el) => {
    const href = $(el).children('div.row').eq(1).find('a').first().attr('href') || null;
    if (href) {
      const full = href.startsWith('http') ? href : 'https://www.tenderdetail.com' + href;
      links.push(full);
    }
    if (links.length >= 3) return false;
  });
  return links;
}

async function scrapeDetail(url) {
  const r = await axios.get(url, {headers:H, timeout:15000, maxRedirects:4});
  const $ = cheerio.load(r.data);
  const record = {};
  $('table tr').each((_, row) => {
    const $tds = $(row).find('td');
    if ($tds.length >= 2) {
      const label = clean($tds.eq(0).text());
      const value = clean($tds.eq(1).text());
      if (!label.startsWith('Download') && label && value && label.length < 80)
        record[label] = value;
    }
  });
  return record;
}

(async () => {
  const workingLinks = await getLinks('https://www.tenderdetail.com/dailytenders/51168758/66998cc5-1b3c-4664-be2e-df09f4bf3c1d');
  const brokenLinks  = await getLinks('https://www.tenderdetail.com/dailytenders/51188939/226e8ac7-3ff2-4dc0-94b7-31dd4de4d76f');

  console.log('=== WORKING URL detail page ===');
  const w = await scrapeDetail(workingLinks[0]);
  console.log('Keys:', Object.keys(w));
  console.log('TDR:', w['TDR'], '| Tender No:', w['Tender No'], '| EMD:', w['EMD']);

  console.log('\n=== BROKEN URL detail page ===');
  const b = await scrapeDetail(brokenLinks[0]);
  console.log('Keys:', Object.keys(b));
  console.log('TDR:', b['TDR'], '| Tender No:', b['Tender No'], '| EMD:', b['EMD']);

  console.log('\n=== Key diff ===');
  const wk = new Set(Object.keys(w));
  const bk = new Set(Object.keys(b));
  const onlyW = [...wk].filter(k => !bk.has(k));
  const onlyB = [...bk].filter(k => !wk.has(k));
  console.log('Only in working:', onlyW);
  console.log('Only in broken:', onlyB);
})().catch(e => console.error('ERROR:', e.message));
