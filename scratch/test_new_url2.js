const axios = require('axios');
const cheerio = require('cheerio');
const H = {'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36','Accept':'text/html'};
const clean = s => (s||'').replace(/\s+/g,' ').trim();

async function main() {
  // Get a viewLink from the new URL
  const listUrl = 'https://www.tenderdetail.com/dailytenders/51188939/226e8ac7-3ff2-4dc0-94b7-31dd4de4d76f';
  const r = await axios.get(listUrl, {headers:H, timeout:15000});
  const $ = cheerio.load(r.data);
  
  // Get first few viewLinks
  const links = [];
  $('div.m-mainTR').each((_, el) => {
    const row2 = $(el).children('div.row').eq(1);
    let href = row2.find('a').first().attr('href') || null;
    if (href && !href.startsWith('http')) href = 'https://www.tenderdetail.com' + href;
    if (href) links.push(href);
    if (links.length >= 3) return false;
  });
  
  console.log('Testing viewLinks:', links.slice(0,3));
  
  // Scrape first detail page
  if (links[0]) {
    console.log('\n--- Detail page ---');
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
    console.log('Keys:', Object.keys(record));
    console.log('TDR:', record['TDR']);
    console.log('Tender No:', record['Tender No']);
    console.log('EMD:', record['EMD']);
    console.log('Tender Value:', record['Tender Value']);
    
    // Test parseDailyDigest sections
    const sections = [];
    let curSection = null, curTenders = [];
    $('p.m-r-government-tenders, div.m-mainTR').each((_, el) => {
      const cls = $(el).attr('class') || '';
      if (cls.includes('m-r-government-tenders')) {
        if (curTenders.length > 0) sections.push({section: curSection || 'All', tenders: curTenders});
        curSection = clean($(el).text());
        curTenders = [];
      } else {
        const row2 = $(el).children('div.row').eq(1);
        let vl = row2.find('a').first().attr('href') || null;
        if (vl && !vl.startsWith('http')) vl = 'https://www.tenderdetail.com' + vl;
        if (vl) curTenders.push({viewLink: vl});
      }
    });
    if (curTenders.length > 0) sections.push({section: curSection || 'All', tenders: curTenders});
    console.log('\nSections found:', sections.length);
    sections.forEach(s => console.log(` - ${s.section}: ${s.tenders.length} tenders`));
  }
}
main().catch(e => console.error('ERROR:', e.message));
