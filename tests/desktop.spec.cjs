const {test,expect,_electron:electron,chromium}=require('@playwright/test');
const path=require('path');
const fs=require('fs');
let app,page,errors;
async function exportFile(selector,destination){
 if(process.env.BROWSER_TEST){const pending=page.waitForEvent('download');await page.click(selector);await (await pending).saveAs(destination);}
 else {await app.evaluate(({BrowserWindow},file)=>{globalThis.exportDone=new Promise((resolve,reject)=>BrowserWindow.getAllWindows()[0].webContents.session.once('will-download',(_event,item)=>{item.setSavePath(file);item.once('done',(_e,status)=>status==='completed'?resolve():reject(new Error(status)));}));},path.resolve(destination));await page.click(selector);await app.evaluate(()=>globalThis.exportDone);}
}
test.beforeEach(async()=>{errors=[];if(process.env.BROWSER_TEST){app=await chromium.launch({channel:'msedge',headless:true});page=await app.newPage({viewport:{width:1440,height:980}});page.on('pageerror',e=>errors.push(e.message));await page.goto('file:///'+path.resolve('src/index.html').replaceAll('\\','/'));}else{app=await electron.launch(process.env.PACKAGED_TEST?{executablePath:path.resolve('dist/v1.1.0/win-unpacked/qPCR studio.exe'),args:[]}:{args:[path.resolve('main.cjs')]});page=await app.firstWindow();page.on('pageerror',e=>errors.push(e.message));}await page.waitForSelector('#exploreDemo');});
test.beforeEach(async()=>{await expect(page.locator('#saveStatus')).toHaveText('Python engine ready',{timeout:90000});});
test.afterEach(async()=>{await app.close();expect(errors).toEqual([]);});
test('demo, figures, project roundtrip, and PNG export',async()=>{
 if(!process.env.PACKAGED_TEST)await page.screenshot({path:'test-results/workspace.png'});
 await page.click('#exploreDemo');await expect(page.locator('#rawTableWrap tbody tr')).toHaveCount(18);
 await page.click('[data-page=setup]');await page.click('#runAnalysis');await expect(page.locator('#resultsWrap tbody tr')).toHaveCount(4);
 await page.click('[data-page=graphs]');await page.selectOption('#palettePreset','ocean');await page.fill('#graphTitle','My experiment');
 for(const type of ['bar','points','box','paired','grouped']){await page.selectOption('#graphType',type);await expect(page.locator('#chart')).toBeVisible();}
 await page.selectOption('#graphMode','multi');await page.locator('[data-graph-gene]').evaluateAll(els=>els.forEach(e=>{if(!e.checked){e.checked=true;e.dispatchEvent(new Event('change'));}}));
 if(!process.env.PACKAGED_TEST)await page.screenshot({path:'test-results/figure-studio.png'});
 await exportFile('#downloadPng','test-results/figure.png');expect(fs.statSync('test-results/figure.png').size).toBeGreaterThan(10000);
 await exportFile('#downloadSvg','test-results/figure.svg');const svg=fs.readFileSync('test-results/figure.svg','utf8');expect(svg).toContain('<svg');expect(svg).toContain('My experiment');
 await exportFile('#downloadPdf','test-results/figure.pdf');expect(fs.readFileSync('test-results/figure.pdf').subarray(0,4).toString()).toBe('%PDF');
 await exportFile('#globalSave','test-results/project.json');const payload=JSON.parse(fs.readFileSync('test-results/project.json'));expect(payload.figure.colors[0]).toBe('#397c96');expect(payload.raw).toHaveLength(18);
 await page.click('[data-page=import]');await page.setInputFiles('#projectInput','test-results/project.json');await page.click('[data-page=graphs]');await expect(page.locator('#graphTitle')).toHaveValue('My experiment');await expect(page.locator('#palettePreset')).toHaveValue('ocean');
});
test('CSV numeric blanks, known fold change, invalidation and missing references',async()=>{
 const csv='Sample,Target,Ct\nControl,GAPDH,20\nControl,GENE,25\nTreatment,GAPDH,20\nTreatment,GENE,23\nTreatment,GENE,\nTreatment,GENE,Undetermined\nMissing,GENE,24';
 await page.setInputFiles('#fileInput',{name:'trial.csv',mimeType:'text/csv',buffer:Buffer.from(csv)});await expect(page.locator('#rawTableWrap tbody tr')).toHaveCount(5);
 await page.click('[data-page=setup]');await page.click('#runAnalysis');
 const treatment=page.locator('#resultsWrap tbody tr').filter({hasText:'Treatment'});await expect(treatment.locator('td').nth(8)).toHaveText('4.000');
 const missing=page.locator('#resultsWrap tbody tr').filter({hasText:'Missing'});await expect(missing.locator('td').nth(8)).toHaveText('—');
 await page.click('[data-page=import]');const row=page.locator('#rawTableWrap tbody tr').filter({has:page.locator('input[value="Treatment"]')}).filter({has:page.locator('input[value="GENE"]')});await row.locator('[data-raw=cq]').fill('24');await row.locator('[data-raw=cq]').press('Tab');
 await page.click('[data-page=results]');await expect(page.locator('#resultsWrap tbody tr').filter({hasText:'Treatment'}).locator('td').nth(8)).toHaveText('2.000');
});
test('offline XLSX import and paired biological test',async()=>{
 const XLSX=require('xlsx');
 for(let i=0;i<3;i++){const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Sample','Target','Ct'],['Control','GAPDH',20],['Control','GENE',25],['Treatment','GAPDH',20],['Treatment','GENE',23+i*.2]]),'Data');await page.setInputFiles('#fileInput',{name:`trial${i}.xlsx`,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:XLSX.write(wb,{type:'buffer',bookType:'xlsx'})});await expect(page.locator('#rawTableWrap tbody tr')).toHaveCount((i+1)*4);}
 await page.click('[data-page=graphs]');await page.click('#runStats');await expect(page.locator('#statsResult')).toContainText('paired biological n = 3');await expect(page.locator('#statsResult')).toContainText('p =');
});
test('dense graphs keep compact export dimensions and diagonal labels',async()=>{
 const rows=['Sample,Target,Ct'];
 for(let c=0;c<4;c++){rows.push(`Condition ${c},GAPDH,20`);for(let g=0;g<12;g++)rows.push(`Condition ${c},Target_${g+1},${25-g*.1-c*.2}`);}
 await page.setInputFiles('#fileInput',{name:'dense.csv',mimeType:'text/csv',buffer:Buffer.from(rows.join('\n'))});
 await expect(page.locator('#rawTableWrap tbody tr')).toHaveCount(52);
 await page.click('[data-page=graphs]');
 await expect(page.locator('#chart')).toBeVisible();await expect(page.locator('#chart')).toHaveAttribute('aria-busy','false');
 const initialWidth=await page.locator('#chart').evaluate(image=>parseFloat(image.style.width));
 await page.selectOption('#graphMode','multi');
 await page.locator('[data-graph-gene]').evaluateAll(els=>els.forEach(e=>{e.checked=true;e.dispatchEvent(new Event('change'));}));
 for(const type of ['grouped','bar']){
  await page.selectOption('#graphType',type);
  await expect(page.locator('#chart')).toHaveAttribute('aria-busy','false');
  const dimensions=await page.locator('#chart').evaluate(image=>({width:parseFloat(image.style.width),available:image.parentElement.clientWidth,angle:image.dataset.labelAngle,engine:image.dataset.engine}));
  expect(dimensions.width).toBe(initialWidth);expect(dimensions.width).toBeLessThanOrEqual(820);expect(dimensions.width).toBeLessThanOrEqual(dimensions.available);
  expect(dimensions.angle).toBe('45');expect(dimensions.engine).toContain('Matplotlib');
  await exportFile('#downloadPng',`test-results/compact-${type}.png`);
  expect(fs.readFileSync(`test-results/compact-${type}.png`).readUInt32BE(16)).toBe(Math.round(initialWidth*3));
 }
});
