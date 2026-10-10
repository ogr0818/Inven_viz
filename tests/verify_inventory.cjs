const {chromium}=require('playwright');
const fs=require('fs'),path=require('path'),assert=require('assert/strict'),{execFileSync}=require('child_process');
const root=path.resolve(__dirname,'..'),out=path.join(root,'verification');
fs.mkdirSync(out,{recursive:true});
const checks=[];const check=(name,fn)=>{fn();checks.push(name)};
(async()=>{
 const browser=await chromium.launch({...(process.env.CHROME_BIN?{executablePath:process.env.CHROME_BIN}:{}),headless:true});
 try{
 const page=await browser.newPage({viewport:{width:1440,height:1000},acceptDownloads:true});
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.goto(process.env.INVENTORY_URL || require('url').pathToFileURL(path.join(root,'inventory.html')).href);
 await page.evaluate(()=>{window.showSaveFilePicker=undefined});
 assert.equal(await page.title(),'住院藥局藥品進耗存');
 assert.equal(await page.locator('#resultCount').innerText(),'367 個品項');
 checks.push('首次開啟採最新月份 11509');
 await page.screenshot({path:path.join(out,'initial.png')});
 const expected=JSON.parse(execFileSync(process.env.PYTHON || 'python3',['-c',`
import sqlite3,json
c=sqlite3.connect('file:summary.db?mode=ro',uri=True)
sets=[]
for start in range(1,10):
 for end in range(start,10):
  rows=c.execute('''WITH selected AS (SELECT * FROM monthly_summary WHERE roc_year=115 AND roc_month BETWEEN ? AND ?), sums AS (SELECT drug_code,SUM(issued_quantity) actual,SUM(inpatient_usage) qty,MAX(roc_month) last_month FROM selected GROUP BY drug_code) SELECT s.drug_code,s.actual,s.qty,m.stock_quantity,11500+s.last_month FROM sums s JOIN selected m ON m.drug_code=s.drug_code AND m.roc_month=s.last_month ORDER BY s.drug_code''',(start,end)).fetchall()
  sets.append({'from':11500+start,'to':11500+end,'rows':rows})
print(json.dumps(sets))
`],{cwd:root,encoding:'utf8'}));
 const actual=await page.evaluate(sets=>sets.map(s=>({...s,rows:Analysis.query(Inventory.rows,{from:s.from,to:s.to,selected:'',sortKey:'id',direction:'asc'}).rows.map(r=>[r.id,r.actual,r.qty,r.stock,r.stockPeriod])})),expected.map(({from,to})=>({from,to})));
 check('45 種連續區間、全品項結果逐值與 SQLite 獨立查詢一致',()=>assert.deepEqual(actual,expected));
 const scenarioResult=await page.evaluate(()=>{
  const row=(year,month,issued,usage,stock)=>({'藥品代碼':'TEST01','藥品名稱':'測試品項','劑型':'O','撥補年份':String(year),'撥補月份':String(month).padStart(2,'0'),'實發量':issued,'住院耗用':usage,'庫存量':stock});
  const data=[row(114,11,1,1,99),row(114,12,5,-2,-3),row(115,1,10,4,0),row(115,2,-1,100,500),row(115,2,0,100,500),row(115,3,20,20,900)];
  const s={from:11412,to:11502,selected:'',sortKey:'id',direction:'asc'};
  const q=Analysis.query(data,s),empty=Analysis.query([data[0],data[4]],s);
  return {rows:q.rows.map(r=>[r.actual,r.qty,r.stock,r.stockPeriod]),missing:q.missing,empty:empty.rows.length};
 });
 check('測試資料：跨年、負實發排除、缺月、零庫存與區間外庫存排除',()=>assert.deepEqual(scenarioResult,{rows:[[15,2,0,11501]],missing:[11502],empty:0}));
 const history=await page.evaluate(()=>{
  const row=(month,name,type)=>({'藥品代碼':'TEST01','藥品名稱':name,'劑型':type,'撥補年份':'115','撥補月份':month,'實發量':5,'住院耗用':1,'庫存量':2});
  const q=Analysis.query([row('02','新藥名','I'),row('01','舊藥名','O')],{from:11501,to:11502,selected:'',sortKey:'id',direction:'asc',types:['I']});
  return {rows:q.rows.map(r=>[r.name,r.type,r.stockPeriod]),detail:q.detail.map(r=>[r['藥品名稱'],r['劑型']])};
 });
 check('歷史名稱與劑型保留，期間統計及劑型篩選使用區間內最新月份',()=>assert.deepEqual(history,{rows:[['新藥名','I',11502]],detail:[['舊藥名','O'],['新藥名','I']]}));
 const embedded=await page.evaluate(()=>Inventory.rows);
 const dbRows=JSON.parse(execFileSync(process.env.PYTHON || 'python3',['-c',`import sqlite3,json;c=sqlite3.connect('file:summary.db?mode=ro',uri=True);print(json.dumps(c.execute('SELECT drug_code,drug_name,issued_quantity,inpatient_usage,stock_quantity,roc_year,roc_month,drug_type FROM monthly_summary ORDER BY roc_year,roc_month,drug_code').fetchall()))`],{cwd:root,encoding:'utf8'}));
 check('內嵌 3,400 筆資料與 DB 逐欄一致，保留負數及零庫存',()=>assert.deepEqual(embedded.map(r=>[r['藥品代碼'],r['藥品名稱'],r['實發量'],r['住院耗用'],r['庫存量'],+r['撥補年份'],+r['撥補月份'],r['劑型']]),dbRows));
 await page.selectOption('#startMonth','1');
 assert.equal(await page.locator('#pending').isVisible(),true);
 assert.equal(await page.locator('#currentState').innerText(),'套用區間：11509');
 await page.click('#apply');
 assert.equal(await page.locator('#resultCount').innerText(),'579 個品項');
 checks.push('年月修改須套用；全期間呈現 579 個品項');
 await page.selectOption('#endMonth','1');await page.selectOption('#startMonth','9');await page.click('#apply');
 assert.match(await page.locator('#feedback').innerText(),/不能晚於/);
 assert.equal(await page.locator('#resultCount').innerText(),'579 個品項');
 await page.selectOption('#startMonth','1');await page.selectOption('#endMonth','9');await page.click('#apply');
 await page.fill('#drugCode','act03i');await page.click('#search');
 assert.equal(await page.locator('#drugCode').inputValue(),'ACT03I');assert.equal(await page.locator('#resultCount').innerText(),'1 個品項');
 await page.fill('#drugCode','zzzzzz');await page.click('#search');
 assert.match(await page.locator('#feedback').innerText(),/沒有配對/);assert.equal(await page.locator('#drugSelect').inputValue(),'ACT03I');
 await page.fill('#drugCode','abc');await page.click('#search');assert.match(await page.locator('#feedback').innerText(),/六碼/);
 checks.push('非法起訖與搜尋失敗保留結果；小寫代碼轉大寫');
 await page.click('#tableTab');assert.equal(await page.locator('#rawDetails').isVisible(),false);assert.equal(await page.locator('#chartSort').isVisible(),false);
 await page.click('#resultsTable th:nth-child(5) button');
 const selection=await page.evaluate(()=>({selected:state.selected,sortKey:state.sortKey,direction:state.direction}));
 await page.click('#chartTab');assert.deepEqual(await page.evaluate(()=>({selected:state.selected,sortKey:state.sortKey,direction:state.direction})),selection);
 checks.push('切換檢視保留品項及排序；表格隱藏明細與排序選單');
 await page.selectOption('#drugSelect','');await page.selectOption('#sort','actual:desc');
 // 驗證月份回退、無資料時保留選擇。
 const absent=await page.evaluate(()=>{const x=result.rows.find(r=>r.stockPeriod<11509);return x.id});
 await page.selectOption('#drugSelect',absent);
 const last=await page.evaluate(()=>result.rows[0].stockPeriod);assert.ok(last<11509);
 await page.selectOption('#startMonth','9');await page.click('#apply');
 assert.equal(await page.locator('#drugSelect').inputValue(),absent);assert.equal(await page.locator('#empty').isVisible(),true);assert.equal(await page.locator('#downloadDetailXlsx').isDisabled(),false);
 await page.selectOption('#startMonth','1');await page.click('#apply');await page.selectOption('#drugSelect','');
 checks.push('庫存僅採區間內最後月份；新區間無所選品項仍保留選擇');
 // 真實下載邊界：表格排序後匯出、明細及完整六頁 PNG ZIP。
 await page.click('#tableTab');await page.click('#resultsTable th:nth-child(5) button');
 let downloadPromise=page.waitForEvent('download');await page.click('#downloadXlsx');let download=await downloadPromise;
 await download.saveAs(path.join(out,'period.xlsx'));
 const expectedTable=await page.evaluate(()=>[columns.map(c=>c[1]),...result.rows.map(r=>columns.map(([k])=>k==='stockPeriod'?String(r[k]):r[k]))]);
 fs.writeFileSync(path.join(out,'expected_period.json'),JSON.stringify(expectedTable));
 await page.click('#chartTab');await page.locator('#rawDetails summary').click();
 downloadPromise=page.waitForEvent('download');await page.click('#downloadDetailXlsx');download=await downloadPromise;
 assert.equal(download.suggestedFilename(),'明細.xlsx');await download.saveAs(path.join(out,'detail.xlsx'));
 fs.writeFileSync(path.join(out,'expected_detail.json'),JSON.stringify(await page.evaluate(()=>[fields,...result.detail.map(r=>fields.map(f=>r[f]))])));
 await page.selectOption('#sort','actual:desc');
 await page.screenshot({path:path.join(out,'full-period-top.png')});
 downloadPromise=page.waitForEvent('download',{timeout:120000});await page.click('#downloadPng');download=await downloadPromise;
 assert.equal(download.suggestedFilename(),'圖表_11501-11509_全.zip');await download.saveAs(path.join(out,'charts.zip'));
 checks.push('完整 579 品項六張圖表 ZIP 實際下載成功');
 await page.selectOption('#drugSelect','ACT03I');
 downloadPromise=page.waitForEvent('download');await page.click('#downloadPng');download=await downloadPromise;
 assert.equal(download.suggestedFilename(),'圖表_11501-11509_ACT03I.png');await download.saveAs(path.join(out,'single.png'));
 checks.push('單品項直接下載 PNG');
 await page.screenshot({path:path.join(out,'single-item.png')});
 // 支援儲存選擇器的控制流程，以測試替身驗證，不聲稱已操作原生對話框。
 await page.evaluate(()=>{window.showSaveFilePicker=async opts=>({name:'自訂表格.xlsx',createWritable:async()=>({write:async blob=>{window.savedSize=blob.size},close:async()=>{},abort:async()=>{}})})});
 await page.click('#tableTab');await page.click('#downloadXlsx');await page.waitForFunction(()=>document.getElementById('tableDownloadStatus').textContent==='已儲存：自訂表格.xlsx');
 assert.ok(await page.evaluate(()=>window.savedSize>0));
 await page.evaluate(()=>{window.showSaveFilePicker=async()=>{throw new DOMException('cancel','AbortError')}});
 await page.click('#downloadXlsx');await page.waitForFunction(()=>document.getElementById('tableDownloadStatus').textContent==='已取消下載。');
 await page.evaluate(()=>{window.showSaveFilePicker=async()=>({name:'fail.xlsx',createWritable:async()=>({write:async()=>{throw new Error('測試寫入失敗')},abort:async()=>{},close:async()=>{}})})});
 await page.click('#downloadXlsx');await page.waitForFunction(()=>document.getElementById('tableDownloadStatus').textContent.includes('測試寫入失敗'));
 checks.push('儲存完成、取消與寫入失敗提示（選擇器測試替身）');
 await page.evaluate(()=>{window.showSaveFilePicker=undefined});
 await page.click('#chartTab');await page.selectOption('#startMonth','1');await page.selectOption('#endMonth','1');await page.click('#apply');
 const negativeBar=await page.evaluate(()=>{const groups=[...document.querySelectorAll('#chart > g')];const line=groups[1].querySelector('line'),bar=groups[1].querySelector('rect');return +bar.getAttribute('x')<+line.getAttribute('x1')});
 assert.ok(negativeBar);await page.screenshot({path:path.join(out,'negative-usage.png')});
 await page.selectOption('#drugSelect','LAC02I');await page.selectOption('#endMonth','5');await page.click('#apply');
 // 負庫存案例以來源資料動態尋找，避免假設指定品項一定存在負值月份。
 const neg=await page.evaluate(()=>Inventory.rows.find(r=>r['庫存量']<0));
 await page.selectOption('#startMonth',String(+neg['撥補月份']));await page.selectOption('#endMonth',String(+neg['撥補月份']));await page.click('#apply');await page.selectOption('#drugSelect',neg['藥品代碼']);
 assert.equal(await page.evaluate(()=>result.rows[0].stock),neg['庫存量']);
 await page.screenshot({path:path.join(out,'negative-stock.png')});
 checks.push('實際負耗用長條朝零點左側，負庫存保留來源值');
 await page.selectOption('#startMonth','1');await page.selectOption('#endMonth','9');await page.click('#apply');await page.selectOption('#drugSelect','ACT03I');
 await page.setViewportSize({width:390,height:844});await page.click('#chartTab');await page.screenshot({path:path.join(out,'mobile.png')});
 assert.ok(await page.evaluate(()=>document.documentElement.scrollWidth<=window.innerWidth));
 checks.push('390px 手機版操作區無整頁水平溢出');
 assert.deepEqual(errors,[]);checks.push('瀏覽器無 JavaScript 執行錯誤');
 fs.writeFileSync(path.join(out,'browser-results.json'),JSON.stringify({checks,errors},null,2));
 console.log(JSON.stringify({checks,errors},null,2));
 }finally{await browser.close()}
})().catch(e=>{console.error(e);process.exit(1)});
