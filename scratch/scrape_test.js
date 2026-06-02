require('dotenv').config();
const axios = require('axios');
const cheerio = require('cheerio');
const H = {'User-Agent':'Mozilla/5.0 (Macintosh) Chrome/120','Accept':'text/html'};
const clean = s => (s||'').replace(/\s+/g,' ').trim();

async function main() {
  // Step 1: fetch listing page
  const listUrl = 'https://www.tenderdetail.com/dailytenders/50824692/d24d3021-bf18-404b-af8d-e66137a47f49';
  console.log('Fetching listing page...');
  const r = await axios.get(listUrl, {headers:H, timeout:15000});
  const $ = cheerio.load(r.data);
  
  const mainTRs = $('div.m-mainTR');
  console.log('div.m-mainTR count:', mainTRs.length);
  
  const links = [];
  mainTRs.each((_, el) => {
    const row2 = $(el).children('div.row').eq(1);
    let href = row2.find('a').first().attr('href') || null;
    if (href && !href.startsWith('http')) href = 'https://www.tenderdetail.com' + href;
    if (href) links.push(href);
  });
  console.log('ViewLinks found:', links.length);
  
  if (links.length === 0) {
    console.log('No viewlinks — checking page structure...');
    console.log('First m-mainTR html:', mainTRs.first().html()?.slice(0,500));
    const tenderLinks = [];
    $('a[href]').each((_, a) => {
      const h = $(a).attr('href')||'';
      if (h.includes('TenderNotice') || h.includes('tender')) tenderLinks.push(h);
    });
    console.log('Any tender links on page:', tenderLinks.slice(0,5));
    return;
  }
  
  // Step 2: scrape first detail page
  console.log('\nScraping detail page:', links[0]);
  const r2 = await axios.get(links[0], {headers:H, timeout:15000});
  const $2 = cheerio.load(r2.data);
  const record = {};
  $2('table tr').each((_, row) => {
    const $tds = $2(row).find('td');
    if ($tds.length >= 2) {
      const label = clean($tds.eq(0).text());
      const value = clean($tds.eq(1).text());
      if (!label.startsWith('Download') && label && value && label.length < 80)
        record[label] = value;
    }
  });
  console.log('Detail page keys:', Object.keys(record));
  console.log('TDR:', record['TDR']);
  console.log('Tender No:', record['Tender No']);
  console.log('Tendering Authority:', record['Tendering Authority']);
  console.log('EMD:', record['EMD']);
  console.log('Tender Value:', record['Tender Value']);
}

main().catch(e => console.error('ERROR:', e.message));
