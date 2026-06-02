require('dotenv').config();
const axios = require('axios');
const cheerio = require('cheerio');
const H = {'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36','Accept':'text/html'};
const clean = s => (s||'').replace(/\s+/g,' ').trim();

async function getSections(label, url) {
  const r = await axios.get(url, {headers:H, timeout:20000});
  const $ = cheerio.load(r.data);
  const sections = [];
  let cur = null, count = 0;
  $('p.m-r-government-tenders, div.m-mainTR').each((_, el) => {
    const cls = $(el).attr('class')||'';
    if (cls.includes('m-r-government-tenders')) {
      if (cur && count > 0) sections.push({name: cur, count});
      cur = clean($(el).text()); count = 0;
    } else {
      count++;
    }
  });
  if (cur && count > 0) sections.push({name: cur, count});
  
  console.log(`\n=== ${label} (${sections.reduce((n,s)=>n+s.count,0)} total) ===`);
  sections.forEach(s => {
    const basename = s.name.replace(/\(\d+\)\s*$/, '').replace(/[\\/:?*[\]]/g, '').trim();
    console.log(`  "${s.name}" → sheet: "${basename}" (${s.count} tenders)`);
  });
  return sections;
}

(async () => {
  await getSections('WORKING', 'https://www.tenderdetail.com/dailytenders/51168758/66998cc5-1b3c-4664-be2e-df09f4bf3c1d');
  await getSections('BROKEN',  'https://www.tenderdetail.com/dailytenders/51188939/226e8ac7-3ff2-4dc0-94b7-31dd4de4d76f');
})().catch(e => console.error('ERROR:', e.message));
