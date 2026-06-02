const axios = require('axios');
const cheerio = require('cheerio');
const H = {'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36','Accept':'text/html'};
const clean = s => (s||'').replace(/\s+/g,' ').trim();

async function main() {
  const url = 'https://www.tenderdetail.com/dailytenders/51188939/226e8ac7-3ff2-4dc0-94b7-31dd4de4d76f';
  console.log('Testing URL:', url);
  const r = await axios.get(url, {headers:H, timeout:15000});
  const $ = cheerio.load(r.data);
  console.log('Page title:', $('title').text().trim());
  console.log('m-mainTR count:', $('div.m-mainTR').length);
  console.log('m-r-government-tenders count:', $('p.m-r-government-tenders').length);
  
  if ($('div.m-mainTR').length > 0) {
    const first = $('div.m-mainTR').first();
    const row2 = first.children('div.row').eq(1);
    const href = row2.find('a').first().attr('href') || '';
    console.log('First viewLink:', href);
    if (!href) {
      console.log('row2 HTML:', row2.html()?.slice(0,400));
      console.log('All anchors in first mainTR:', first.find('a').map((_,a) => $(a).attr('href')).get());
    }
  } else {
    console.log('No m-mainTR found. Body text:', $('body').text().replace(/\s+/g,' ').slice(0,300));
  }
}
main().catch(e => console.error('ERROR:', e.message));
