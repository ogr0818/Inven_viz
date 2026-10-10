const {chromium}=require('playwright');
const fs=require('fs'),path=require('path'),assert=require('assert/strict'),{execFileSync}=require('child_process');
const root=path.resolve(__dirname,'..'),out=path.join(root,'verification'),checks=[];
(async()=>{
 const browser=await chromium.launch({...(process.env.CHROME_BIN?{executablePath:process.env.CHROME_BIN}:{}),headless:true});
 try{
 const p=await browser.newPage({viewport:{width:1440,height:1050},acceptDownloads:true,offline:!process.env.INVENTORY_URL});
 const errors=[];p.on('pageerror',e=>errors.push(e.message));
 await p.goto(process.env.INVENTORY_URL || require('url').pathToFileURL(path.join(root,'inventory.html')).href);
 await p.evaluate(()=>{window.showSaveFilePicker=undefined});
 assert.equal(await p.locator('#minIssued').inputValue(),'');assert.equal(await p.locator('#maxIssued').inputValue(),'');assert.equal(await p.locator('[name="drugType"]:checked').count(),0);
 assert.equal(await p.locator('#apply').innerText(),'套用條件');assert.equal(await p.locator('#pending').isVisible(),false);
 checks.push('初始上下限為空、四劑型未勾選，未篩選結果與原版一致');
 // SQL 對照全部 16 種劑型組合與包含／單邊上下限，共 144 組。
 const expected=JSON.parse(execFileSync(process.env.PYTHON || 'python3',['-c',`
import sqlite3,json,itertools
c=sqlite3.connect('file:summary.db?mode=ro',uri=True)
sets=[]
for size in range(5):
 for types in itertools.combinations('EIOS',size):
  for lo,hi in [(None,None),(1,None),(None,100),(100,100),(100,1000),(1000,20000),(99999999,None),(None,1),(180,180)]:
   rows=c.execute('''WITH sums AS (SELECT drug_code,SUM(issued_quantity) actual,SUM(inpatient_usage) qty,MAX(roc_month) last_month FROM monthly_summary WHERE roc_year=115 GROUP BY drug_code) SELECT s.drug_code,s.actual,s.qty,m.stock_quantity,11500+s.last_month,m.drug_type FROM sums s JOIN monthly_summary m ON m.drug_code=s.drug_code AND m.roc_year=115 AND m.roc_month=s.last_month WHERE (? IS NULL OR s.actual>=?) AND (? IS NULL OR s.actual<=?) ORDER BY s.drug_code''',(lo,lo,hi,hi)).fetchall()
   rows=[r for r in rows if not types or r[5] in types]
   ids={r[0] for r in rows}
   detail=c.execute('SELECT drug_code,roc_month,issued_quantity,inpatient_usage,stock_quantity FROM monthly_summary WHERE roc_year=115 ORDER BY roc_month,drug_code').fetchall()
   detail=[r for r in detail if r[0] in ids]
   sets.append({'minIssued':lo,'maxIssued':hi,'types':types,'rows':rows,'detail':detail})
print(json.dumps(sets))
`],{cwd:root,encoding:'utf8',maxBuffer:20*1024*1024}));
 const actual=await p.evaluate(sets=>sets.map(s=>{const q=Analysis.query(Inventory.rows,{...s,from:11501,to:11509,selected:'',sortKey:'id',direction:'asc'});return {...s,rows:q.rows.map(r=>[r.id,r.actual,r.qty,r.stock,r.stockPeriod,r.type]),detail:q.detail.map(r=>[r['藥品代碼'],+r['撥補月份'],r['實發量'],r['住院耗用'],r['庫存量']])}}),expected.map(({minIssued,maxIssued,types})=>({minIssued,maxIssued,types})));
 assert.deepEqual(actual,expected);checks.push('144 組範圍／劑型組合與獨立 SQL 一致，含邊界與完整逐月資料');
 await p.selectOption('#startMonth','1');await p.fill('#minIssued','100');await p.fill('#maxIssued','1000');await p.check('[name="drugType"][value="I"]');await p.check('[name="drugType"][value="O"]');
 assert.equal(await p.locator('#currentState').innerText(),'套用區間：11509');assert.equal(await p.locator('#resultCount').innerText(),'367 個品項');
 assert.equal(await p.locator('#pending').isVisible(),true);assert.match(await p.locator('#pending').innerText(),/請點擊「套用條件」/);assert.match(await p.locator('#apply').getAttribute('class'),/needs-apply/);
 await p.screenshot({path:path.join(out,'filters-pending.png')});
 await p.click('#apply');assert.equal(await p.locator('#pending').isVisible(),false);assert.match(await p.locator('#appliedStatus').innerText(),/100～1,000.*I、O/);
 const target=expected.find(s=>s.minIssued===100&&s.maxIssued===1000&&s.types.join('')==='IO');
 const ids=await p.locator('#drugSelect option').evaluateAll(options=>options.map(o=>o.value).filter(Boolean));
 assert.deepEqual([...ids].sort(),target.rows.map(r=>r[0]).sort());assert.equal(await p.locator('#chart .drug').count(),ids.length);
 await p.click('#tableTab');assert.equal(await p.locator('#resultsTable tbody tr').count(),ids.length);await p.click('#chartTab');
 checks.push('年月、範圍與劑型共同套用；輸入中保留結果並顯示待套用提示，選單／圖表／表格同步');
 await p.selectOption('#drugSelect','ACT03I');await p.fill('#minIssued','181');assert.equal(await p.locator('#resultCount').innerText(),'1 個品項');await p.click('#apply');
 assert.equal(await p.locator('#drugSelect').inputValue(),'ACT03I');assert.equal(await p.locator('#empty').isVisible(),true);assert.match(await p.locator('#empty').innerText(),/不符合目前篩選條件/);
 let downloadCount=0;p.on('download',()=>downloadCount++);
 await p.evaluate(()=>{window.pickerCalls=0;window.showSaveFilePicker=async()=>{window.pickerCalls++;throw Error('空結果不可開啟選擇器')}});
 await p.click('#downloadPng');assert.match(await p.locator('#downloadStatus').innerText(),/沒有符合條件.*無法下載/);
 await p.locator('#rawDetails summary').click();await p.click('#downloadDetailXlsx');assert.match(await p.locator('#detailDownloadStatus').innerText(),/無法下載/);
 await p.click('#tableTab');await p.click('#downloadXlsx');assert.match(await p.locator('#tableDownloadStatus').innerText(),/無法下載/);
 assert.equal(downloadCount,0);assert.equal(await p.evaluate(()=>window.pickerCalls),0);
 await p.selectOption('#drugSelect','');assert.equal(await p.locator('#drugSelect').inputValue(),'');assert.ok((await p.locator('#resultsTable tbody tr').count())>0);assert.match(await p.locator('#appliedStatus').innerText(),/181～1,000.*I、O/);
 checks.push('單品項不符合時保留選擇並提示；三種下載均攔截且不開啟儲存選擇器，清除品項不清除篩選');
 const old=await p.locator('#resultsTable tbody').innerText();
 for(const invalid of ['0','-1','1.5','abc','1e3','1,000','9007199254740992']){
  await p.fill('#minIssued',invalid);await p.click('#apply');assert.equal(await p.locator('#minIssued').getAttribute('aria-invalid'),'true');assert.ok((await p.locator('#minIssuedError').innerText()).length>0);assert.equal(await p.locator('#resultsTable tbody').innerText(),old);
 }
 await p.fill('#minIssued','500');await p.fill('#maxIssued','100');await p.click('#apply');assert.match(await p.locator('#maxIssuedError').innerText(),/不可小於/);assert.equal(await p.locator('#resultsTable tbody').innerText(),old);
 checks.push('0、負數、小數、文字、科學記號、逗號、超大整數與上下限顛倒皆報錯且保留結果');
 // 套用空白與劑型全不勾選可恢復全量；修改後復原亦撤除待套用提示。
 await p.fill('#minIssued','');await p.fill('#maxIssued','');await p.uncheck('[name="drugType"][value="I"]');await p.uncheck('[name="drugType"][value="O"]');await p.click('#apply');assert.equal(await p.locator('#resultsTable tbody tr').count(),579);
 await p.check('[name="drugType"][value="S"]');assert.equal(await p.locator('#pending').isVisible(),true);await p.uncheck('[name="drugType"][value="S"]');assert.equal(await p.locator('#pending').isVisible(),false);
 checks.push('清空上下限且全部未勾選恢復全量；草稿改回原值即撤除待套用提示');
 // 有效篩選後實際下載三種輸出。
 await p.fill('#minIssued','100');await p.fill('#maxIssued','1000');await p.check('[name="drugType"][value="I"]');await p.check('[name="drugType"][value="O"]');await p.click('#apply');
 await p.evaluate(()=>{window.showSaveFilePicker=undefined});
 await p.click('#resultsTable th:nth-child(5) button');
 let promise=p.waitForEvent('download');await p.click('#downloadXlsx');let d=await promise;await d.saveAs(path.join(out,'filtered-period.xlsx'));
 fs.writeFileSync(path.join(out,'expected_filtered_period.json'),JSON.stringify(await p.evaluate(()=>[columns.map(c=>c[1]),...result.rows.map(r=>columns.map(([k])=>k==='stockPeriod'?String(r[k]):r[k]))])));
 await p.click('#chartTab');if(!await p.locator('#rawDetails').evaluate(el=>el.open))await p.locator('#rawDetails summary').click();
 promise=p.waitForEvent('download');await p.click('#downloadDetailXlsx');d=await promise;await d.saveAs(path.join(out,'filtered-detail.xlsx'));
 fs.writeFileSync(path.join(out,'expected_filtered_detail.json'),JSON.stringify(await p.evaluate(()=>[fields,...result.detail.map(r=>fields.map(f=>r[f]))])));
 await p.selectOption('#sort','actual:desc');
 const count=await p.evaluate(()=>result.rows.length);
 promise=p.waitForEvent('download',{timeout:120000});await p.click('#downloadPng');d=await promise;
 await d.saveAs(path.join(out,count>100?'filtered-charts.zip':'filtered-chart.png'));
 fs.writeFileSync(path.join(out,'filtered-count.json'),JSON.stringify({count,detail:await p.evaluate(()=>result.detail.length)}));
 checks.push('有效篩選的圖表、期間統計與月明細均真實下載成功');
 await p.evaluate(()=>scrollTo(0,0));await p.screenshot({path:path.join(out,'filters-applied.png')});
 await p.setViewportSize({width:390,height:844});await p.fill('#maxIssued','2000');await p.evaluate(()=>scrollTo(0,0));await p.screenshot({path:path.join(out,'filters-mobile.png')});
 assert.ok(await p.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));assert.deepEqual(errors,[]);
 checks.push('手機篩選操作與待套用提示可見，無整頁水平溢出；無 JavaScript 錯誤');
 fs.writeFileSync(path.join(out,'filter-results.json'),JSON.stringify({checks,errors},null,2));console.log(JSON.stringify({checks,errors},null,2));
 }finally{await browser.close()}
})().catch(e=>{console.error(e);process.exit(1)});
