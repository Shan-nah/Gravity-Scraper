// Test end-to-end: scrape 3 tenders from the new URL and run buildExcel
require('dotenv').config();
const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');
const fs = require('fs');

const H = {'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36','Accept':'text/html'};
const clean = s => (s||'').replace(/\s+/g,' ').trim();

async function scrapeDetail(viewLink) {
  try {
    const {data} = await axios.get(viewLink, {headers:H, timeout:15000, maxRedirects:4});
    const $ = cheerio.load(data);
    const record = {};
    $('table tr').each((_,row) => {
      const $tds = $(row).find('td');
      if ($tds.length >= 2) {
        const label = clean($tds.eq(0).text());
        const value = clean($tds.eq(1).text());
        if (!label.startsWith('Download') && label && value && label.length < 80)
          record[label] = value;
      }
    });
    return {
      'Company':'','Important':false,'Filled Date':'','Filled By':'','Bid Status':'',
      'TDR': record['TDR']||'N/A',
      'Tender No': record['Tender No']||record['Tender ID']||'N/A',
      'Tendering Authority': record['Tendering Authority']||record['Company Name']||'N/A',
      'Tender Brief': record['Tender Brief']||'N/A',
      'City': record['City']||'N/A','State': record['State']||'N/A',
      'Document Fees': record['Document Fees']||'N/A',
      'EMD': parseFloat(String(record['EMD']||'').replace(/[^\d.]/g,''))||'N/A',
      'EMD Exempt?':'No','EMD Exemption':'N/A',
      'Tender Value': parseFloat(String(record['Tender Value']||'').replace(/[^\d.]/g,''))||'N/A',
      'Tender Type': record['Tender Type']||'N/A',
      'Bidding Type': record['Bidding Type']||'N/A',
      'Competition Type': record['Competition Type']||'N/A',
      'Publish Date': record['Publish Date']||'N/A',
      'Last Date of Bid Submission': record['Last Date of Bid Submission']||'N/A',
      'Tender Opening Date': record['Tender Opening Date']||'N/A',
      'Address': record['Address']||'N/A',
      'Information Source': record['Information Source']||'N/A',
      'View Link': viewLink,'Additional Details':'N/A'
    };
  } catch(e) {
    return {'TDR':'ERR','Tender No':'N/A','Company':'','Important':false,'Filled Date':'',
      'Filled By':'','Bid Status':'','Tendering Authority':'N/A','Tender Brief':e.message,
      'City':'N/A','State':'N/A','Document Fees':'N/A','EMD':'N/A',
      'EMD Exempt?':'N/A','EMD Exemption':'N/A','Tender Value':'N/A',
      'Tender Type':'N/A','Bidding Type':'N/A','Competition Type':'N/A',
      'Publish Date':'N/A','Last Date of Bid Submission':'N/A',
      'Tender Opening Date':'N/A','Address':'N/A','Information Source':'N/A',
      'View Link':viewLink,'Additional Details':'N/A'};
  }
}

async function main() {
  // Get first 3 tenders from new URL per section
  const r = await axios.get('https://www.tenderdetail.com/dailytenders/51188939/226e8ac7-3ff2-4dc0-94b7-31dd4de4d76f',{headers:H,timeout:15000});
  const $ = cheerio.load(r.data);
  
  const sections = [];
  let cur = null, curT = [];
  $('p.m-r-government-tenders, div.m-mainTR').each((_,el) => {
    const cls = $(el).attr('class')||'';
    if(cls.includes('m-r-government-tenders')) {
      if(curT.length>0) sections.push({section:cur||'All',tenders:curT.slice(0,2)});
      cur = clean($(el).text()); curT=[];
    } else {
      const row2 = $(el).children('div.row').eq(1);
      let href = row2.find('a').first().attr('href')||null;
      if(href&&!href.startsWith('http')) href='https://www.tenderdetail.com'+href;
      if(href&&curT.length<2) curT.push({viewLink:href});
    }
  });
  if(curT.length>0) sections.push({section:cur||'All',tenders:curT.slice(0,2)});
  
  console.log(`Sections: ${sections.length}, total tenders: ${sections.reduce((n,s)=>n+s.tenders.length,0)}`);
  
  // Scrape tenders
  const enriched = [];
  for(const sec of sections) {
    const records = await Promise.all(sec.tenders.map(t => scrapeDetail(t.viewLink)));
    enriched.push({section:sec.section, tenders:records});
    console.log(`Section "${sec.section}": ${records.length} records, first TDR: ${records[0]?.TDR}`);
  }
  
  // Try buildExcel - we need the real buildExcel from server.js
  // Since we can't easily import it, let's at least verify allRows
  const COLS_KEYS = ['Company','Important','Section','TDR','Tender No','Tendering Authority','Tender Brief','City','State','Document Fees','EMD','EMD Exempt?','EMD Exemption','Tender Value','Tender Type','Bidding Type','Competition Type','Publish Date','Last Date of Bid Submission','Tender Opening Date','Address','Information Source','View Link','Additional Details'];
  const allRows = enriched.flatMap(s => s.tenders.map(t => ({Section:s.section,...t})));
  console.log(`\nallRows count: ${allRows.length}`);
  console.log('First row keys present in COLS_KEYS:');
  const row0 = allRows[0];
  COLS_KEYS.forEach(k => {
    const v = row0[k];
    console.log(`  ${k}: ${JSON.stringify(v === undefined ? 'MISSING' : v).slice(0,50)}`);
  });
}
main().catch(e => console.error('FATAL:', e.message, e.stack?.split('\n').slice(0,3).join(' ')));
