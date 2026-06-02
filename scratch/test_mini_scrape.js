require('dotenv').config();
const axios = require('axios');
const cheerio = require('cheerio');
const H = {'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36','Accept':'text/html'};
const clean = s => (s||'').replace(/\s+/g,' ').trim();

// Simulate what the server does end-to-end with 2 tenders per section from the broken URL
async function main() {
  const listUrl = 'https://www.tenderdetail.com/dailytenders/51188939/226e8ac7-3ff2-4dc0-94b7-31dd4de4d76f';
  const r = await axios.get(listUrl, {headers:H, timeout:20000});
  const $ = cheerio.load(r.data);
  
  // Replicate parseDailyDigest
  const sections = [];
  let curSection = null, curTenders = [];
  $('p.m-r-government-tenders, div.m-mainTR').each((_, el) => {
    const $el = $(el);
    const cls = $el.attr('class')||'';
    if (cls.includes('m-r-government-tenders')) {
      if (curTenders.length > 0) sections.push({section:curSection||'All',tenders:curTenders});
      curSection = clean($el.text()); curTenders = [];
    } else {
      const row2 = $el.children('div.row').eq(1);
      let href = row2.find('a').first().attr('href')||null;
      if (href && !href.startsWith('http')) href = 'https://www.tenderdetail.com'+href;
      const desc = $el.children('div.row').eq(0).find('.col-md-12').first();
      let tenderId = 'N/A';
      desc.find('p.m-td-brief').each((_,p) => {
        const t = clean($(p).text());
        if (/TDR:\d+/i.test(t) && tenderId==='N/A') {
          const m = t.match(/TDR:(\d+)/i); if(m) tenderId=m[1];
        }
      });
      if (href && curTenders.length < 2) curTenders.push({tenderId, viewLink:href});
    }
  });
  if (curTenders.length > 0) sections.push({section:curSection||'All',tenders:curTenders});
  
  console.log(`Sections: ${sections.length}`);
  sections.forEach(s => console.log(`  "${s.section}": ${s.tenders.length} tenders`));
  
  // Now call the actual /test-excel endpoint with this data via POST to server
  // Instead, let's just post to the scrape endpoint to trigger the full flow
  // and see what the server console says
  
  // Simulate what buildExcel would get - look at allRows construction
  const enriched = sections.map(s => ({
    section: s.section,
    tenders: s.tenders.map(t => ({
      'Company':'','Important':false,'Filled Date':'','Filled By':'','Bid Status':'',
      'TDR':t.tenderId,'Tender No':'N/A','Tendering Authority':'N/A',
      'Tender Brief':'Test','City':'N/A','State':'N/A','Document Fees':'N/A',
      'EMD':1000,'EMD Exempt?':'No','EMD Exemption':'N/A','Tender Value':10000,
      'Tender Type':'Open','Bidding Type':'Item Rate','Competition Type':'Open',
      'Publish Date':'01-06-2026','Last Date of Bid Submission':'20-06-2026',
      'Tender Opening Date':'21-06-2026','Address':'N/A','Information Source':'N/A',
      'View Link':t.viewLink,'Additional Details':'N/A'
    }))
  }));
  
  const COLS = [
    {key:'Company',label:'Company',width:16},{key:'Important',label:'Important',width:12},
    {key:'Filled Date',label:'Filled Date',width:16},{key:'Filled By',label:'Filled By',width:18},
    {key:'Bid Status',label:'Bid Status',width:14},{key:'TDR',label:'TDR',width:14},
    {key:'Tender No',label:'Tender No',width:26},{key:'Tendering Authority',label:'Tendering Authority',width:36},
    {key:'Tender Brief',label:'Tender Brief',width:62},{key:'City',label:'City',width:16},
    {key:'State',label:'State',width:18},{key:'Document Fees',label:'Document Fees',width:16},
    {key:'EMD',label:'EMD (₹)',width:18},{key:'EMD Exempt?',label:'EMD Exempt?',width:13},
    {key:'EMD Exemption',label:'EMD Exemption',width:44},{key:'Tender Value',label:'Tender Value (₹)',width:18},
    {key:'Tender Type',label:'Tender Type',width:16},{key:'Bidding Type',label:'Bidding Type',width:16},
    {key:'Competition Type',label:'Competition Type',width:18},{key:'Publish Date',label:'Publish Date',width:14},
    {key:'Last Date of Bid Submission',label:'Last Date of Bid Submission',width:26},
    {key:'Tender Opening Date',label:'Tender Opening Date',width:22},
    {key:'Address',label:'Address',width:34},{key:'Information Source',label:'Information Source',width:30},
    {key:'View Link',label:'View Link',width:60},{key:'Additional Details',label:'Additional Details',width:48},
  ];
  const MASTER_INPUT_KEYS = new Set(['Company','Important']);
  const ALL_INPUT_KEYS = new Set(['Company','Important','Filled Date','Filled By','Bid Status']);
  const finalCols = COLS.filter(c => c.key !== 'Bid Document Details');
  const masterCols = [
    ...finalCols.filter(c => MASTER_INPUT_KEYS.has(c.key)),
    {key:'Section',label:'Section',width:28},
    ...finalCols.filter(c => !ALL_INPUT_KEYS.has(c.key)),
  ];
  
  const allRows = enriched.flatMap(s => s.tenders.map(t => ({Section:s.section,...t})));
  
  console.log(`\nallRows: ${allRows.length}`);
  console.log('masterCols count:', masterCols.length);
  
  // Check each row has all masterCols keys
  let missingKeyCount = 0;
  allRows.forEach((r, i) => {
    masterCols.forEach(c => {
      if (r[c.key] === undefined) {
        console.log(`Row ${i}: MISSING key "${c.key}"`);
        missingKeyCount++;
      }
    });
  });
  if (missingKeyCount === 0) console.log('All rows have all keys ✓');
  
  // Check section names vs filterValues
  console.log('\nSection filter checks:');
  sections.forEach(s => {
    const basename = s.section.replace(/\(\d+\)\s*$/, '').replace(/[\\/:?*[\]]/g, '').trim();
    const sectionVals = allRows.filter(r => r.Section === s.section).map(r => r.Section);
    const matches = allRows.filter(r => r.Section.includes(basename));
    console.log(`  Sheet "${basename}": SEARCH finds ${matches.length} of ${sectionVals.length} expected`);
  });
}
main().catch(e => console.error('FATAL:', e.message));
